"""
图片场景自动分类工具
使用Ollama本地视觉模型(LLaVA)识别图片场景内容，按场景分类到不同文件夹
"""

import os
import shutil
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

try:
    from PIL import Image
    from PIL.ExifTags import TAGS
except ImportError:
    Image = None
    TAGS = None

try:
    from ollama import Client
except ImportError:
    raise ImportError(
        "请先安装 ollama 库: pip install ollama\n"
        "同时需要安装 Ollama 应用: https://ollama.com/download"
    )

from logger_setup import setup_logger


class ImageClassifier:
    """图片场景分类器 - 使用本地Ollama视觉模型"""

    # 支持的图片格式
    IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp', '.webp'}

    # 场景描述用于AI判断
    SCENE_DESCRIPTIONS = {
        "会议": "会议场景 - 如会议现场、会议记录、会议PPT、讨论照片、参会人员等",
        "检查": "检查场景 - 如安全检查、质量检查、整改通知、问题记录、违规通知等",
        "培训": "培训场景 - 如培训照片、培训记录、培训现场、培训材料等",
        "活动": "活动场景 - 如团建活动、公司活动、集体活动、庆祝活动等",
        "验收": "验收场景 - 如工程验收、项目验收、交付验收、验收报告等",
        "施工": "施工场景 - 如施工现场、工地照片、施工过程、施工记录等",
    }

    # 未分类类别
    UNCLASSIFIED = "未分类"

    # Ollama模型配置
    DEFAULT_MODEL = "llava:7b"  # 默认使用LLaVA 7B模型
    FALLBACK_MODEL = "moondream"  # 备选模型

    # AI分类提示词 - 针对LLaVA优化
    CLASSIFICATION_PROMPT = """You are an image classification expert. Please analyze this image and determine which scene it belongs to:

1. **会议** - Meeting scenes, including: meeting sites, meeting minutes/records, meeting PPTs, discussion photos, attendees, seminars, symposiums, etc.

2. **检查** - Inspection scenes, including: safety inspections, quality inspections, rectification notices, problem records, hazard investigations, violation notices, etc.

3. **培训** - Training scenes, including: training photos, training records, training sites, training materials, safety training, skill training, etc.

4. **活动** - Activity scenes, including: team building activities, company events, group activities, celebration events, parties, ceremonies, etc.

5. **验收** - Acceptance scenes, including: engineering acceptance, project acceptance, delivery acceptance, acceptance reports, final inspections, etc.

6. **施工** - Construction scenes, including: construction sites, work site photos, construction processes, construction records, building work, etc.

7. **未分类** - Images that do not belong to the above categories.

Please respond with ONLY ONE word: 会议, 检查, 培训, 活动, 验收, 施工, or 未分类

Analyze the image now:"""

    def __init__(
        self,
        source_dir: str,
        output_dir: Optional[str] = None,
        model: str = DEFAULT_MODEL,
        host: str = "http://localhost:11434"
    ):
        """
        初始化分类器

        Args:
            source_dir: 源图片文件夹路径
            output_dir: 输出文件夹路径，默认为源目录下的"分类文件夹"
            model: Ollama模型名称 (默认: llava:7b)
            host: Ollama服务地址 (默认: http://localhost:11434)
        """
        self.source_dir = Path(source_dir)
        if output_dir is None:
            self.output_dir = self.source_dir.parent / "分类文件夹"
        else:
            self.output_dir = Path(output_dir)

        self.model = model
        self.logger = setup_logger("image_classifier")

        # 初始化Ollama客户端
        try:
            self.client = Client(host=host)
            # 测试连接
            self.client.list()
            self.logger.info(f"已连接到Ollama服务: {host}")
            self.logger.info(f"使用模型: {model}")
        except Exception as e:
            raise ConnectionError(
                f"无法连接到Ollama服务 ({host})。\n"
                f"请确保: 1) Ollama已安装并运行  2) 已下载模型: ollama pull {model}\n"
                f"下载地址: https://ollama.com/download\n"
                f"错误详情: {e}"
            )

        # 统计信息 - 动态初始化以支持所有场景类别
        self.stats: Dict[str, List[str]] = {
            scene: [] for scene in list(self.SCENE_DESCRIPTIONS.keys()) + [self.UNCLASSIFIED]
        }

    def get_image_files(self) -> List[Path]:
        """获取源目录中所有图片文件"""
        image_files = []
        for file_path in self.source_dir.rglob("*"):
            if file_path.is_file() and file_path.suffix.lower() in self.IMAGE_EXTENSIONS:
                image_files.append(file_path)
        return image_files

    def extract_image_date(self, image_path: Path) -> Optional[datetime]:
        """
        从图片EXIF数据中提取拍摄日期
        优先级：DateTimeOriginal > DateTime > 文件修改时间

        Args:
            image_path: 图片文件路径

        Returns:
            拍摄日期的datetime对象，如果无法提取则返回文件修改时间
        """
        # 如果PIL不可用，直接使用文件修改时间
        if Image is None:
            return datetime.fromtimestamp(image_path.stat().st_mtime)

        try:
            with Image.open(image_path) as img:
                exif_data = img._getexif()

                if exif_data is None:
                    # 无EXIF数据，使用文件修改时间
                    return datetime.fromtimestamp(image_path.stat().st_mtime)

                # 优先级1: DateTimeOriginal（原始拍摄时间）
                for tag_id, value in exif_data.items():
                    tag = TAGS.get(tag_id, tag_id)
                    if tag == "DateTimeOriginal":
                        try:
                            return datetime.strptime(value, "%Y:%m:%d %H:%M:%S")
                        except (ValueError, TypeError):
                            continue

                # 优先级2: DateTime（修改时间）
                for tag_id, value in exif_data.items():
                    tag = TAGS.get(tag_id, tag_id)
                    if tag == "DateTime":
                        try:
                            return datetime.strptime(value, "%Y:%m:%d %H:%M:%S")
                        except (ValueError, TypeError):
                            continue

                # 无法从EXIF获取日期，使用文件修改时间
                return datetime.fromtimestamp(image_path.stat().st_mtime)

        except Exception as e:
            self.logger.debug(f"读取EXIF数据失败 {image_path.name}: {e}")
            return datetime.fromtimestamp(image_path.stat().st_mtime)

    def classify_by_ai(self, image_path: Path) -> str:
        """
        使用Ollama视觉模型(LLaVA)分析图片并分类

        Args:
            image_path: 图片文件路径

        Returns:
            场景类别: "会议", "检查", 或 "未分类"
        """
        try:
            # Ollama直接读取图片文件，需要绝对路径
            response = self.client.chat(
                model=self.model,
                messages=[{
                    'role': 'user',
                    'content': self.CLASSIFICATION_PROMPT,
                    'images': [str(image_path.resolve())]  # 传递绝对路径
                }],
                options={
                    'temperature': 0.1,  # 降低温度以获得更稳定的分类
                    'num_predict': 10,   # 限制输出长度
                }
            )

            # 解析AI返回结果
            result_text = response['message']['content'].strip()

            self.logger.debug(f"AI分类结果: {result_text}")

            # 标准化结果文本
            result_text_lower = result_text.lower()

            # 根据AI返回判断类别 - 优先匹配中文，其次英文
            if "会议" in result_text or "meeting" in result_text_lower:
                return "会议"
            elif "检查" in result_text or "inspection" in result_text_lower:
                return "检查"
            elif "培训" in result_text or "training" in result_text_lower:
                return "培训"
            elif "活动" in result_text or "activity" in result_text_lower or "event" in result_text_lower:
                return "活动"
            elif "验收" in result_text or "acceptance" in result_text_lower:
                return "验收"
            elif "施工" in result_text or "construction" in result_text_lower:
                return "施工"
            else:
                return self.UNCLASSIFIED

        except Exception as e:
            self.logger.warning(f"AI识别失败 {image_path.name}: {e}")
            return self.UNCLASSIFIED

    def ensure_output_dirs(self) -> None:
        """确保输出目录结构存在"""
        for scene in list(self.SCENE_DESCRIPTIONS.keys()) + [self.UNCLASSIFIED]:
            scene_dir = self.output_dir / scene
            scene_dir.mkdir(parents=True, exist_ok=True)

    def move_file(self, src: Path, scene: str) -> Path:
        """
        移动文件到对应场景文件夹并重命名为"场景名+日期"格式

        Args:
            src: 源文件路径
            scene: 场景类别

        Returns:
            目标文件路径
        """
        dest_dir = self.output_dir / scene

        # 提取图片日期并格式化为 YYYYMMDD
        image_date = self.extract_image_date(src)
        date_str = image_date.strftime("%Y%m%d")

        # 生成新文件名: 场景名+日期.扩展名 (如: 会议20260102.jpg)
        new_filename = f"{scene}{date_str}{src.suffix}"
        dest_path = dest_dir / new_filename

        # 处理重名文件 - 添加序号 _1, _2, _3, ...
        counter = 1
        while dest_path.exists():
            dest_path = dest_dir / f"{scene}{date_str}_{counter}{src.suffix}"
            counter += 1

        shutil.move(str(src), str(dest_path))
        return dest_path

    def generate_report(self) -> Path:
        """
        生成分类报告

        Returns:
            报告文件路径
        """
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_path = self.output_dir / f"分类报告_{timestamp}.txt"

        with open(report_path, "w", encoding="utf-8") as f:
            f.write("=" * 60 + "\n")
            f.write("图片场景分类报告 (Ollama + LLaVA)\n")
            f.write("=" * 60 + "\n\n")
            f.write(f"处理时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"使用模型: {self.model}\n")
            f.write(f"源目录: {self.source_dir}\n")
            f.write(f"输出目录: {self.output_dir}\n\n")

            total = 0
            for scene, files in self.stats.items():
                count = len(files)
                total += count
                f.write(f"\n【{scene}】共 {count} 个文件\n")
                f.write("-" * 40 + "\n")
                for file_name in files:
                    f.write(f"  - {file_name}\n")

            f.write("\n" + "=" * 60 + "\n")
            f.write(f"总计处理: {total} 个文件\n")
            f.write("=" * 60 + "\n")

        return report_path

    def process(self) -> Tuple[int, int]:
        """
        执行分类处理

        Returns:
            (成功数量, 失败数量)
        """
        self.logger.info(f"开始处理图片分类")
        self.logger.info(f"源目录: {self.source_dir}")

        # 确保输出目录存在
        self.ensure_output_dirs()

        # 获取所有图片文件
        image_files = self.get_image_files()
        self.logger.info(f"找到 {len(image_files)} 个图片文件")

        if not image_files:
            self.logger.warning("未找到任何图片文件")
            return 0, 0

        # 处理每个文件
        success_count = 0
        for i, image_path in enumerate(image_files, 1):
            self.logger.info(f"[{i}/{len(image_files)}] 处理: {image_path.name}")

            # AI视觉识别分类
            scene = self.classify_by_ai(image_path)
            self.logger.info(f"  AI识别场景: {scene}")

            # 移动文件
            try:
                dest_path = self.move_file(image_path, scene)
                self.stats[scene].append(image_path.name)
                self.logger.info(f"  已移动到: {dest_path.relative_to(self.output_dir.parent)}")
                success_count += 1
            except Exception as e:
                self.logger.error(f"  移动文件失败: {e}")

        # 生成报告
        report_path = self.generate_report()
        self.logger.info(f"分类完成，报告已生成: {report_path}")

        return success_count, len(image_files) - success_count


def main():
    """命令行入口"""
    import argparse

    parser = argparse.ArgumentParser(
        description="图片场景自动分类工具 (使用Ollama本地视觉模型 - 免费、无API限制)"
    )
    parser.add_argument("source_dir", help="源图片文件夹路径")
    parser.add_argument("-o", "--output", help="输出文件夹路径（默认: 源目录父目录/分类文件夹）")
    parser.add_argument(
        "-m", "--model",
        default="llava:7b",
        help="Ollama模型名称 (默认: llava:7b, 备选: moondream)"
    )
    parser.add_argument(
        "-H", "--host",
        default="http://localhost:11434",
        help="Ollama服务地址 (默认: http://localhost:11434)"
    )

    args = parser.parse_args()

    # 验证源目录
    source_path = Path(args.source_dir)
    if not source_path.exists():
        print(f"错误: 源目录不存在: {args.source_dir}")
        return 1

    if not source_path.is_dir():
        print(f"错误: 源路径不是目录: {args.source_dir}")
        return 1

    # 执行分类
    try:
        classifier = ImageClassifier(
            str(source_path),
            args.output,
            model=args.model,
            host=args.host
        )
        success, failed = classifier.process()

        print(f"\n处理完成: 成功 {success}, 失败 {failed}")
        return 0
    except ValueError as e:
        print(f"错误: {e}")
        return 1
    except Exception as e:
        print(f"处理失败: {e}")
        return 1


if __name__ == "__main__":
    exit(main())
