import os
import sys


def get_file_size(file_path):
    """获取文件大小，处理权限错误"""
    try:
        return os.path.getsize(file_path)
    except (OSError, IOError, PermissionError):
        return 0


def scan_path_and_sort_files(target_path):
    """
    扫描指定路径中的所有最基层文件，并按大小排序

    Args:
        target_path: 要扫描的路径，可以是驱动器或文件夹

    Returns:
        sorted_files: 按文件大小排序的文件列表
    """
    file_list = []

    # 检查路径是否存在
    if not os.path.exists(target_path):
        print(f"错误: 路径 '{target_path}' 不存在!")
        return []

    # 检查是否是文件而非目录
    if os.path.isfile(target_path):
        print(f"注意: '{target_path}' 是一个文件，不是目录。将扫描其所在目录。")
        target_path = os.path.dirname(target_path)

    print(f"开始扫描 '{target_path}' ...")
    print("这可能需要一些时间，请耐心等待...\n")

    # 计数器
    file_count = 0
    error_count = 0

    try:
        for root, dirs, files in os.walk(target_path):
            # 处理当前目录下的所有文件（最基层文件）
            for file in files:
                file_path = os.path.join(root, file)
                file_count += 1

                # 每扫描10000个文件显示一次进度
                if file_count % 10000 == 0:
                    print(f"已扫描 {file_count} 个文件...")

                try:
                    file_size = get_file_size(file_path)
                    file_list.append((file_path, file_size))
                except Exception:
                    error_count += 1
                    continue

    except KeyboardInterrupt:
        print("\n用户中断扫描")
        return []
    except Exception as e:
        print(f"扫描过程中出现错误: {e}")
        return []

    print(f"\n扫描完成！")
    print(f"总共扫描文件: {file_count} 个")
    print(f"成功获取大小的文件: {len(file_list)} 个")
    print(f"无法访问的文件: {error_count} 个")

    # 按文件大小降序排序（从大到小）
    print("\n正在按文件大小排序...")
    sorted_files = sorted(file_list, key=lambda x: x[1], reverse=True)

    return sorted_files


def format_file_size(size_bytes):
    """将文件大小格式化为易读的格式"""
    if size_bytes == 0:
        return "0 B"

    size_names = ["B", "KB", "MB", "GB", "TB"]
    i = 0
    while size_bytes >= 1024 and i < len(size_names) - 1:
        size_bytes /= 1024.0
        i += 1

    return f"{size_bytes:.2f} {size_names[i]}"


def display_results(sorted_files, top_n=50):
    """显示排序结果"""
    if not sorted_files:
        print("没有找到文件或扫描被中断")
        return

    print(f"\n{'=' * 80}")
    print(f"前 {top_n} 个最大的文件:")
    print(f"{'排名':<6} {'文件大小':<12} {'文件路径'}")
    print(f"{'-' * 80}")

    for i, (file_path, size) in enumerate(sorted_files[:top_n], 1):
        formatted_size = format_file_size(size)
        # 如果路径太长，进行截断显示
        display_path = file_path if len(file_path) <= 100 else file_path[:67] + "..."
        print(f"{i:<6} {formatted_size:<12} {display_path}")

    # 显示统计信息
    total_size = sum(size for _, size in sorted_files)
    avg_size = total_size / len(sorted_files) if sorted_files else 0

    print(f"\n统计信息:")
    print(f"总文件大小: {format_file_size(total_size)}")
    print(f"平均文件大小: {format_file_size(avg_size)}")
    print(f"最大文件: {format_file_size(sorted_files[0][1])}")
    print(f"最小文件: {format_file_size(sorted_files[-1][1])}")


def save_to_file(sorted_files, target_name, filename=None):
    """将结果保存到文件"""
    if filename is None:
        # 创建安全的文件名
        safe_name = "".join(c for c in target_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
        filename = f"{safe_name}_文件大小报告.txt"

    try:
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(f"'{target_name}' 文件大小排序报告\n")
            f.write("=" * 50 + "\n\n")

            for i, (file_path, size) in enumerate(sorted_files, 1):
                formatted_size = format_file_size(size)
                f.write(f"{i:4d}. {formatted_size:>10} - {file_path}\n")

        print(f"\n完整结果已保存到: {filename}")
    except Exception as e:
        print(f"保存文件时出错: {e}")


def main():
    """主函数"""
    print("文件大小扫描和排序程序")
    print("=" * 50)
    print("提示:")
    print("- 输入驱动器字母扫描整个驱动器 (例如: C)")
    print("- 输入完整路径扫描特定文件夹 (例如: C:\\Users\\YourName\\Documents)")
    print("- 输入相对路径扫描当前目录下的文件夹 (例如: .\\Downloads)")
    print("=" * 50)

    target = input("请输入要扫描的驱动器或文件夹路径: ").strip()

    # 处理简单的驱动器字母输入
    if len(target) == 1 and target.isalpha():
        target_path = target + ":\\"
        display_name = target + "盘"
    else:
        target_path = target
        display_name = f"'{target}'"

    print(f"\n准备扫描: {display_name}")

    # 扫描文件
    sorted_files = scan_path_and_sort_files(target_path)

    if sorted_files:
        # 显示前50个最大的文件
        display_results(sorted_files, top_n=50)

        # 询问是否保存完整结果
        response = input("\n是否保存完整结果到文件? (y/n): ").lower()
        if response in ['y', 'yes']:
            save_to_file(sorted_files, display_name)

    input("\n按回车键退出...")


if __name__ == "__main__":
    main()