import xml.etree.ElementTree as ET
from pathlib import Path
import logging

LOGGER = logging.getLogger(__name__)

class XMLDataExtractor:
    def __init__(self):
        self.extracted_data = []

    def extract_from_file(self, xml_path: Path):
        if xml_path.suffix.lower() != '.xml':
            return
        
        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()
            
            import re
            raw_text = xml_path.read_text(encoding='utf-8', errors='ignore')
            
            total_pagar = None
            # 1. Buscar <Text name="TotalPagar">...</Text> que a veces viene dentro de un CDATA
            match = re.search(r'<[a-zA-Z0-9_:]*Text\s+name=[\'"]TotalPagar[\'"][^>]*>\s*([^<]+)\s*</', raw_text, re.IGNORECASE)
            if match:
                total_pagar = match.group(1).strip()
            else:
                # 2. Buscar tag <TotalPagar> directo
                match = re.search(r'<[a-zA-Z0-9_:]*TotalPagar[^>]*>\s*([^<]+)\s*</', raw_text, re.IGNORECASE)
                if match:
                    total_pagar = match.group(1).strip()
                else:
                    # 3. Fallback a PayableAmount
                    match = re.search(r'<[a-zA-Z0-9_:]*PayableAmount[^>]*>\s*([^<]+)\s*</', raw_text, re.IGNORECASE)
                    if match:
                        total_pagar = match.group(1).strip()
            
            # Si a pesar de esto sigue nulo, intentamos con ElementTree clásico
            if total_pagar is None:
                tree = ET.parse(xml_path)
                root = tree.getroot()
                for elem in root.iter():
                    tag_name = elem.tag.split('}')[-1]
                    if tag_name in ['TotalPagar', 'PayableAmount']:
                        total_pagar = elem.text
                        break
                    if tag_name == 'Text' and elem.attrib.get('name') == 'TotalPagar':
                        total_pagar = elem.text
                        break
            
            # Crear lista con otra lista dentro, lista interior vacia por el momento
            self.extracted_data.append([total_pagar, []])
        except Exception as e:
            LOGGER.warning(f"Error parsing XML for total: {e}")
            self.extracted_data.append([None, []])

    def get_data(self):
        return self.extracted_data
