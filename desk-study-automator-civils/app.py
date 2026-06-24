import viktor as vkt
from pathlib import Path
import base64
import os


class Parametrization(vkt.Parametrization):
    pass


class Controller(vkt.Controller):
    parametrization = Parametrization

    @vkt.WebView('FRA Generator')
    def get_web_view(self, params, **kwargs):
        html_path = Path(__file__).parent / 'files' / 'FRA_Generator.html'
        docx_path = Path(__file__).parent / 'files' / 'FRA.docx'

        html_content = html_path.read_text(encoding='utf-8')
        docx_b64 = base64.b64encode(docx_path.read_bytes()).decode('ascii')
        html_content = html_content.replace('%%TEMPLATE_B64%%', docx_b64, 1)
        
        # Inject Anthropic API key from environment variable
        anthropic_api_key = os.getenv('ANTHROPIC_API_KEY', '')
        html_content = html_content.replace('%%ANTHROPIC_API_KEY%%', anthropic_api_key)

        return vkt.WebResult(html=vkt.File.from_data(html_content.encode('utf-8')))
