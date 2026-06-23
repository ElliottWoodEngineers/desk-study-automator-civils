import viktor as vkt
from pathlib import Path


class Parametrization(vkt.Parametrization):
    pass


class Controller(vkt.Controller):
    parametrization = Parametrization

    @vkt.WebView('FRA Generator')
    def get_web_view(self, params, **kwargs):
        html_path = Path(__file__).parent / 'files' / 'FRA_Generator.html'
        return vkt.WebResult.from_path(html_path)
