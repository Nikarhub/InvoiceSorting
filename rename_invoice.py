import os
import re
import pdfplumber
import sys  # 必须导入 sys 模块

def extract_invoice_info(file_path):
    filename = os.path.basename(file_path)
    
    try:
        with pdfplumber.open(file_path) as pdf:
            first_page = pdf.pages[0]
            text = first_page.extract_text() or ""
            # 打印原始文本以调试
            # print(f"原始文本: {text}")
            # print("-"*40)
            
            # 1. 日期处理 (纯数字)
            date_match = re.search(r'开票日期[:：]\s*(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日', text)
            if date_match:
                year = date_match.group(1)
                month = date_match.group(2).zfill(2)
                day = date_match.group(3).zfill(2)
                date_str = f"{year}{month}{day}"
            else:
                date_str = "未知日期"

            # 2. 发票号
            invoice_num = "未知发票号"
            labeled_num = re.search(r'发票号码[:：]\s*(\d{20}|\d{8,})', text)
            if labeled_num:
                invoice_num = labeled_num.group(1)
            else:
                raw_num = re.search(r'(?<!\d)(\d{20})(?!\d)', text)
                if raw_num:
                    invoice_num = raw_num.group(1)

            # 3. 订单号 (宽容模式)
            order_num = "无订单号"
            order_pattern = r'(?:订\s*单\s*号|订\s*单\s*编\s*号|Order\s*No)[:：]\s*([A-Za-z0-9]+)'
            order_match = re.search(order_pattern, text, re.IGNORECASE)
            
            if order_match:
                val = order_match.group(1)
                if len(val) > 5:
                    order_num = val

            print(f"[{filename}] -> 订单号:{order_num} | 发票号码:{invoice_num} | 开票日期:{date_str}")
            return order_num, invoice_num, date_str

    except Exception as e:
        print(f"❌ 解析异常 {filename}: {e}")
        return None, None, None

def batch_rename_invoices(folder_path):
    print(f"正在扫描文件夹: {folder_path} ...\n" + "-"*40)
    
    files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
    if not files:
        print("⚠️  文件夹是空的，请放入 PDF 发票。")
        return

    success_count = 0
    for filename in files:
        full_path = os.path.join(folder_path, filename)
        order_no, invoice_no, invoice_date = extract_invoice_info(full_path)
        
        if invoice_no != "未知发票号":
            new_name = f"{order_no}-{invoice_no}-{invoice_date}.pdf"
            new_name = new_name.replace("/", "").replace(":", "").strip()
            new_full_path = os.path.join(folder_path, new_name)

            if full_path == new_full_path:
                print("   -> 跳过: 文件名已正确")
            else:
                if os.path.exists(new_full_path):
                    base, ext = os.path.splitext(new_name)
                    i = 1
                    while os.path.exists(os.path.join(folder_path, f"{base}_{i}{ext}")):
                        i += 1
                    new_full_path = os.path.join(folder_path, f"{base}_{i}{ext}")
                
                try:
                    os.rename(full_path, new_full_path)
                    print(f"   -> ✅ 成功重命名")
                    success_count += 1
                except Exception as e:
                    print(f"   -> ❌ 失败: {e}")
        else:
            print("   -> ⚠️ 跳过: 无法识别")
        print("-" * 40)

    print(f"\n共处理 {len(files)} 个文件，成功修改 {success_count} 个。")

if __name__ == '__main__':
    # --- 核心修改：正确获取 EXE 所在路径 ---
    if getattr(sys, 'frozen', False):
        # 如果是打包后的 EXE 运行，获取 EXE 文件的路径
        base_dir = os.path.dirname(sys.executable)
    else:
        # 如果是脚本运行，获取脚本文件的路径
        base_dir = os.path.dirname(os.path.abspath(__file__))
    
    target_folder = os.path.join(base_dir, 'invoice')
    
    print(f"当前工作目录: {base_dir}")
    print(f"目标发票目录: {target_folder}")
    print("="*40)

    if not os.path.exists(target_folder):
        try:
            os.makedirs(target_folder)
            print(f"✅ 已在当前目录下创建 'invoice' 文件夹！")
            print(f"路径: {target_folder}")
            print("请将发票放入该文件夹后，重新运行程序。")
        except Exception as e:
            print(f"❌ 创建文件夹失败: {e}")
    else:
        batch_rename_invoices(target_folder)
    
    print("\n" + "="*40)
    input("程序执行完毕，请按【回车键】退出...")