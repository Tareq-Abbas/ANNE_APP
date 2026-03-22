from kivy.lang.builder import Builder
from kivy.uix.screenmanager import Screen
from kivy.clock import Clock
from kivymd.app import MDApp





class LoginScreen(Screen):
    pass


class ScannerScreen(Screen):

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        Clock.schedule_once(self._after_init)
        # self.ids.zbarcam_id.ids.xcamera.play=True

    def _after_init(self, dt):
        """
        Binds `ZBarCam.on_symbols()` event.
        """
        zbarcam = self.ids.zbarcam_id
        zbarcam.bind(symbols=self.on_symbols)

    def on_symbols(self, zbarcam, symbols):
        """
        Loads the first symbol data to the `QRFoundScreen.data_property`.
        """
        # going from symbols found to no symbols found state would also
        # trigger `on_symbols`
        if not symbols:
            return

        symbol = symbols[0]
        data = symbol.data.decode('utf8')
        print(data)
        self.manager.get_screen('qr').ids.data.text = data
        self.manager.transition.direction = 'left'
        self.manager.current = 'qr'

    def on_leave(self):
        zbarcam = self.ids.zbarcam_id
        zbarcam.stop()

class QRScreen(Screen):
    pass


class DemoApp(MDApp):
    def build(self):
        # screen =Screen()

        self.title = 'Demeter'
        self.theme_cls.theme_style = "Dark"
        self.theme_cls.primary_palette = "DeepPurple"

        self.help = Builder.load_file('main.kv')
        return self.help


DemoApp().run()