import json
import subprocess
import sys
import re
import tempfile
import os

def run_officecli(args):
    result = subprocess.run(["officecli"] + args + ["--json"], capture_output=True, text=True)
    if result.returncode != 0:
        return None
    try:
        return json.loads(result.stdout)
    except json.JSONDecodeError:
        return None

def main(file_path):
    print(f"[*] Analyzing {os.path.basename(file_path)}...")
    body_data = run_officecli(["get", file_path, "/body", "--depth", "10"])
    if not body_data or "data" not in body_data:
        print("[!] Error: Could not read document body.")
        return

    batch_commands = []
    header_found = False
    last_was_subject = False # Để bắt các dòng nối tiếp của V/v:
    
    # 1. Page Setup
    batch_commands.append({
        "command": "set", "path": "/",
        "props": {
            "pageWidth": "11906", "pageHeight": "16838",
            "marginTop": "1134", "marginBottom": "1134",
            "marginLeft": "1701", "marginRight": "850"
        }
    })

    def process_node(node, in_table=False):
        nonlocal header_found, last_was_subject
        node_type = node.get("type")
        node_path = node.get("path")
        
        if node_type == "table":
            batch_commands.append({
                "command": "set", "path": node_path,
                "props": {"width": "100%", "layout": "auto"}
            })

        elif node_type == "paragraph":
            text = node.get("text", "").strip()
            if not text and not in_table: return
            
            upper_text = text.upper()
            
            # 1. HEADER DETECTION (Chuyển text thành bảng 3 cột nếu chưa có)
            if not header_found and ("TẬP ĐOÀN" in upper_text) and \
               ("CỘNG HÒA" in upper_text or "ĐỘC LẬP" in upper_text or "HẠNH PHÚC" in upper_text):
                header_found = True
                process_header_to_table(node_path, text)
                return

            if in_table:
                # CẢI TIẾN: Tất cả nội dung trong bảng mặc định căn giữa theo yêu cầu
                props = {
                    "font": "Times New Roman", "size": "12pt",
                    "lineSpacing": "1.0x", "spaceBefore": "0pt", "spaceAfter": "0pt",
                    "bold": "false", "italic": "false", "alignment": "center"
                }
                
                # Header của bảng (dòng 1) hoặc các từ khóa quan trọng
                if "[tr[1]]" in node_path or node.get("format", {}).get("bold") or \
                   any(x in upper_text for x in ["TẬP ĐOÀN", "CỘNG HÒA", "ĐỘC LẬP"]):
                    props["bold"] = "true"
                
                # Ngày tháng trong bảng Header thường in nghiêng
                if any(x in upper_text for x in ["TP. HCM", "NGÀY", "THÁNG", "NĂM"]) and "/body/tbl[1]" in node_path:
                    props["italic"] = "true"
                    props["size"] = "13pt"

                batch_commands.append({"command": "set", "path": node_path, "props": props})
            else:
                props = {
                    "font": "Times New Roman", "size": "13pt", "alignment": "justify",
                    "lineSpacing": "1.3x", "spaceBefore": "3pt", "spaceAfter": "6pt",
                    "firstLineIndent": "1.25cm", "widowControl": "false",
                    "bold": "false", "italic": "false"
                }
                
                # Nhận diện TỜ TRÌNH
                if re.match(r'^(TỜ TRÌNH|BÁO CÁO|THÔNG BÁO|QUYẾT ĐỊNH|BIÊN BẢN)', upper_text) and len(text) < 100:
                    props.update({"size": "16pt", "alignment": "center", "bold": "true", "spaceBefore": "12pt", "spaceAfter": "12pt", "firstLineIndent": "0cm"})
                    last_was_subject = False
                
                # CẢI TIẾN: Nhận diện V/v: và các dòng nối tiếp (không bắt đầu bằng số/Kính gửi)
                elif re.match(r'^(V/v:|Về việc:|V/V:)', text, re.IGNORECASE) or \
                     (last_was_subject and len(text) < 150 and not re.match(r'^([0-9]|Kính gửi|Trân trọng)', text)):
                    props.update({"alignment": "center", "bold": "true", "italic": "true", "spaceBefore": "0pt", "spaceAfter": "12pt", "firstLineIndent": "0cm"})
                    last_was_subject = True # Đánh dấu để dòng tiếp theo có thể là continuation
                
                # Kính gửi:
                elif text.startswith("Kính gửi:") or text.startswith("Kính gửi "):
                    props.update({"alignment": "left", "bold": "true", "spaceBefore": "12pt", "spaceAfter": "6pt", "firstLineIndent": "0cm"})
                    last_was_subject = False
                
                # Đề mục lớn
                elif re.match(r'^([0-9]+(\.[0-9]+)*|I[VXLCDM]*|V[IXLCDM]*|X[IXLCDM]*)\.?\s', text):
                    props.update({"alignment": "left", "bold": "true", "spaceBefore": "6pt", "spaceAfter": "6pt", "firstLineIndent": "0cm"})
                    last_was_subject = False
                
                # Tiểu mục
                elif re.match(r'^([a-z]\)\s|\-\s)', text):
                    props.update({"firstLineIndent": "1.25cm", "bold": "false"})
                    last_was_subject = False
                
                # Kết luận
                elif "Trân trọng kính trình" in text or "Trân trọng!" in text:
                    props.update({"italic": "true", "firstLineIndent": "1.25cm"})
                    last_was_subject = False
                else:
                    last_was_subject = False

                batch_commands.append({"command": "set", "path": node_path, "props": props})

        for child in node.get("children", []):
            process_node(child, in_table=in_table or node_type == "table")

    def process_header_to_table(p_path, text):
        clean_text = re.sub(r'(\s{4,}|\t+)', '|', text)
        parts = [p.strip() for p in clean_text.split('|') if p.strip()]
        
        info = {
            'company': next((s for s in parts if "TẬP ĐOÀN" in s.upper()), "TẬP ĐOÀN HOÀN CẦU"),
            'motto': next((s for s in parts if "CỘNG HÒA" in s.upper()), "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"),
            'greeting': next((s for s in parts if any(x in s.upper() for x in ["ĐỘC LẬP", "HẠNH PHÚC"])), "Độc lập – Tự do – Hạnh phúc"),
            'number': next((s for s in parts if "SỐ:" in s.upper()), "Số: ...../2026/TTr-HCG"),
            'date': next((s for s in parts if any(x in s.upper() for x in ["TP. HCM", "NGÀY", "NĂM"])), "Tp. HCM, ngày ...... tháng ...... năm 2026")
        }

        batch_commands.append({
            "command": "add", "path": "/body", "type": "table", "index": 0,
            "props": {"rows": "3", "cols": "3", "width": "100%", "border.all": "none"}
        })
        
        cells = [
            ("/body/tbl[1]/tr[1]/tc[1]/p[1]", info['company'], {"bold": "true", "size": "12pt"}),
            ("/body/tbl[1]/tr[1]/tc[3]/p[1]", info['motto'], {"bold": "true", "size": "12pt"}),
            ("/body/tbl[1]/tr[2]/tc[1]/p[1]", info['number'], {"size": "12pt"}),
            ("/body/tbl[1]/tr[2]/tc[3]/p[1]", info['greeting'], {"bold": "true", "size": "13pt"}),
            ("/body/tbl[1]/tr[3]/tc[3]/p[1]", info['date'], {"italic": "true", "size": "13pt"})
        ]
        
        for cell_p_path, cell_text, extra in cells:
            p_props = {"text": cell_text, "alignment": "center", "font": "Times New Roman", "bold": "false", "italic": "false"}
            p_props.update(extra)
            batch_commands.append({"command": "set", "path": cell_p_path, "props": p_props})
            
        batch_commands.append({"command": "remove", "path": p_path})

    process_node(body_data["data"])

    if batch_commands:
        # Normalize props values to string to ensure JSON compatibility
        for cmd in batch_commands:
            if "props" in cmd:
                for k, v in cmd["props"].items():
                    cmd["props"][k] = str(v).lower() if isinstance(v, bool) else str(v)

        print(f"[*] Applying {len(batch_commands)} updates in one pass...")
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            json.dump(batch_commands, f)
            temp_name = f.name
        
        try:
            subprocess.run(["officecli", "batch", file_path, "--input", temp_name], check=True)
            print("[+] Formatting completed successfully.")
        except subprocess.CalledProcessError:
            print("[!] Error: Batch formatting failed.")
        finally:
            if os.path.exists(temp_name):
                os.remove(temp_name)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 format_hoancau_v2.py <file.docx>")
    else:
        main(sys.argv[1])
