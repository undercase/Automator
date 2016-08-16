import kivy

from kivy.app import App
from kivy.uix.label import Label

class AutomatorApp(App):

    def build(self):
        return Label(text='Hello world')

if __name__ == '__main__':
    AutomatorApp().run()
