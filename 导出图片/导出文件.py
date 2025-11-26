import os
import win32com.client as win32

def export_visio_to_jpeg(input_folder, output_folder):
    # 创建输出目录
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 打开 Visio 应用
    visio = win32.Dispatch("Visio.Application")
    visio.Visible = False  # 不显示界面，加速执行

    # 遍历 input 文件夹
    for filename in os.listdir(input_folder):
        if filename.lower().endswith((".vsd", ".vsdx")):
            visio_path = os.path.join(input_folder, filename)
            print(f"正在处理：{visio_path}")

            try:
                # 打开文档
                doc = visio.Documents.Open(visio_path)

                # 遍历每页
                for page in doc.Pages:
                    # 输出文件名：原文件名 + 页名
                    output_name = f"{os.path.splitext(filename)[0]}_{page.Name}.jpg"
                    output_path = os.path.join(output_folder, output_name)

                    # Export 方式直接导出 JPEG
                    # Visio 会自动进行矢量渲染后导成高质量位图
                    page.Export(output_path)

                    print(f"  导出 -> {output_path}")

                doc.Close()

            except Exception as e:
                print(f"文件处理失败：{filename}，错误：{e}")

    visio.Quit()
    print("\n全部文件导出完成！")


# =============================
# 修改成你的路径
# =============================

input_folder = r"E:\BaiduSyncdisk\专利\构型\设计\0号机\专利二改\专利成品图\专利成品图"
output_folder = r"E:\BaiduSyncdisk\专利\构型\设计\0号机\专利二改\专利成品图\专利成品图\批量图片"

export_visio_to_jpeg(input_folder, output_folder)