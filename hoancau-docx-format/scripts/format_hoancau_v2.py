import json
import subprocess
import sys
import re
import tempfile
import os
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

def run_officecli(args):
    """Executes officecli command and returns JSON output if successful."""
    try:
        result = subprocess.run(["officecli"] + args + ["--json"], capture_output=True, text=True, check=True)
        return json.loads(result.stdout)
    except subprocess.CalledProcessError as e:
        logging.error(f"OfficeCLI Error: {e.stderr}")
        return None
    except json.JSONDecodeError:
        logging.error("Failed to decode JSON from OfficeCLI output.")
        return None

class HoanCauFormatter:
    def __init__(self, file_path):
        self.file_path = file_path
        self.batch_commands = []
        self.paragraphs = []
        self.tables = []

    def load_data(self):
        """Loads paragraph and table data from the document."""
        p_query = run_officecli(["query", self.file_path, "paragraph"])
        if p_query:
            self.paragraphs = p_query["data"]["Results"]
        
        t_query = run_officecli(["query", self.file_path, "table"])
        if t_query:
            self.tables = t_query["data"]["Results"]

    def add_page_setup(self):
        """Adds A4 page setup and standard margins."""
        self.batch_commands.append({
            "command": "set",
            "path": "/",
            "props": {
                "pageWidth": "11906",
                "pageHeight": "16838",
                "marginTop": "1134",
                "marginBottom": "1134",
                "marginLeft": "1701",
                "marginRight": "850"
            }
        })

    def process_paragraphs(self):
        """Iterates through paragraphs and applies formatting rules."""
        header_para_path = None
        header_info = {}

        for p in self.paragraphs:
            p_path = p["path"]
            if "/tbl[" in p_path: continue # Skip paragraphs inside tables

            p_data = run_officecli(["get", self.file_path, p_path])
            if not p_data: continue
            
            p_text = p_data["data"].get("text", "").strip()
            
            # 1. Header Detection
            if "TẬP ĐOÀN" in p_text and "CỘNG HÒA" in p_text:
                header_para_path = p_path
                parts = re.split(r'\s{3,}', p_text)
                header_info['company'] = next((s for s in parts if "TẬP ĐOÀN" in s), "TẬP ĐOÀN HOÀN CẦU")
                header_info['motto'] = next((s for s in parts if "CỘNG HÒA" in s), "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
                header_info['greeting'] = next((s for s in parts if "Độc lập" in s), "Độc lập – Tự do – Hạnh phúc")
                header_info['number'] = next((s for s in parts if "Số:" in s), "Số: ...../2026/TTr-HCG")
                header_info['date'] = next((s for s in parts if any(x in s for x in ["Tp. HCM", "ngày", "năm"])), "Tp. HCM, ngày ...... tháng ...... năm 2026")
                continue

            # 2. Cleanup: Remove empty paragraphs at the end of the doc (optional but cleaner)
            if not p_text and p == self.paragraphs[-1]:
                self.batch_commands.append({"command": "remove", "path": p_path})
                continue

            # 3. Default Body Props
            props = {
                "font": "Times New Roman", "size": "13pt", "alignment": "justify",
                "lineSpacing": "1.3x", "spaceBefore": "3pt", "spaceAfter": "6pt",
                "firstLineIndent": "1.25cm", "widowControl": "false"
            }

            # 4. Heading Detection (Numeric & Roman)
            if re.match(r'^([0-9]+(\.[0-9]+)*)\.', p_text) or re.match(r'^([IVXLCDM]+)\.', p_text):
                props.update({"alignment": "left", "bold": "true", "spaceBefore": "6pt", "spaceAfter": "6pt", "firstLineIndent": "0cm"})

            # 5. Special Patterns
            if "TỜ TRÌNH" in p_text and len(p_text) < 30:
                props.update({"size": "16pt", "alignment": "center", "bold": "true", "spaceBefore": "12pt", "spaceAfter": "12pt", "firstLineIndent": "0cm"})
            elif p_text.startswith("V/v:"):
                props.update({"alignment": "center", "bold": "true", "italic": "true", "spaceBefore": "0pt", "spaceAfter": "12pt", "firstLineIndent": "0cm"})
            elif p_text.startswith("Kính gửi:"):
                props.update({"alignment": "left", "bold": "true", "spaceBefore": "12pt", "spaceAfter": "6pt", "firstLineIndent": "0cm"})

            self.batch_commands.append({"command": "set", "path": p_path, "props": props})

        self.apply_header_table(header_para_path, header_info)

    def apply_header_table(self, path, info):
        """Constructs the hidden header table."""
        if not path: return
        
        self.batch_commands.append({
            "command": "add", "path": "/", "type": "table", "before": "/body/p[1]",
            "props": {"rows": "3", "cols": "3", "width": "100%", "border.all": "none"}
        })
        
        # Mapping table cells (1-based index)
        cells = [
            ("/body/tbl[1]/tr[1]/tc[1]", info.get('company'), {"bold": "true", "size": "12pt"}),
            ("/body/tbl[1]/tr[1]/tc[3]", info.get('motto'), {"bold": "true", "size": "12pt"}),
            ("/body/tbl[1]/tr[2]/tc[1]", info.get('number'), {"size": "12pt"}),
            ("/body/tbl[1]/tr[2]/tc[3]", info.get('greeting'), {"bold": "true", "size": "13pt"}),
            ("/body/tbl[1]/tr[3]/tc[3]", info.get('date'), {"italic": "true", "size": "13pt"})
        ]
        
        for p_path, text, extra in cells:
            p_props = {"text": text, "alignment": "center", "font": "Times New Roman"}
            p_props.update(extra)
            self.batch_commands.append({"command": "set", "path": p_path, "props": p_props})
            
        self.batch_commands.append({"command": "remove", "path": path})

    def process_tables(self):
        """Ensures all tables fit within page margins."""
        for tbl in self.tables:
            self.batch_commands.append({
                "command": "set", "path": tbl["path"],
                "props": {"width": "100%", "layout": "auto"}
            })

    def run_batch(self):
        """Writes batch commands to a temp file and executes them."""
        if not self.batch_commands:
            logging.info("No changes to apply.")
            return

        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            json.dump(self.batch_commands, f)
            temp_name = f.name

        try:
            logging.info(f"Executing {len(self.batch_commands)} formatting commands...")
            subprocess.run(["officecli", "batch", self.file_path, "--input", temp_name], check=True)
            logging.info("Formatting complete.")
        finally:
            os.remove(temp_name)

def main():
    if len(sys.argv) < 2:
        print("Usage: python format_hoancau_v2.py <file.docx>")
        sys.exit(1)

    formatter = HoanCauFormatter(sys.argv[1])
    formatter.load_data()
    formatter.add_page_setup()
    formatter.process_paragraphs()
    formatter.process_tables()
    formatter.run_batch()

if __name__ == "__main__":
    main()
