import os
import sys

import kivy
import PyPDF2
import docx

from kivy.app import App
from kivy.uix.label import Label
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.popup import Popup

class Automator(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)

    def parse_contract(self, file_path, phrase):
        extension = os.path.splitext(file_path)[1]

        if extension == '.pdf':
            self.parse_pdf(file_path, phrase)
        elif extension == '.docx':
            self.parse_docx(file_path, phrase)
        else:
            popup = Popup(title='Error',
                          content=Label(text='The contract must be either a word document or a pdf.'),
                          size_hint=(None, None),
                          size=(400,200))
            popup.open()

    def parse_docx(self, file_path, phrase):
        output = open(file_path.replace('docx', 'csv'), 'w')
        output.write('SECTION, PARAGRAPH, LINE\n')
        document = docx.Document(file_path)

        current_section = None
        current_paragraph = None

        for paragraph in document.paragraphs:
            line = paragraph.text.strip()
            if line.lower().split(' ')[0] == 'section':
                current_section = line.split(' ')[0] + ' ' + line.split(' ')[1]
                print('Current Section: %s' % current_section)
            elif len(line) >= 2 and line[1] == '.':
                line = line.split(' ')[0]
                for character_number in range(len(line)):
                    character = line[character_number]
                    if character_number % 2 != 0:
                        if character != '.':
                            break
                    else:
                        if not (character.isdigit() or character.isupper()):
                            break
                else:
                    if line[-1] != '.':
                        current_paragraph = line
                        print('Current Paragraph: %s' % current_paragraph)
            elif phrase.lower() in line.lower():
                if current_section is not None and current_paragraph is not None:
                    print('Phrase Found')
                    output.write('"%s", "%s", "%s"\n' % (current_section, current_paragraph, line))
        output.close() 
        popup = Popup(title='Completed',
                content=Label(text='The processing has been completed.\nResults are stored in:\n%s' % file_path.replace('pdf', 'csv')),
                size_hint=(None, None),
                size=(600, 200))
        popup.open()


    def parse_pdf(self, file_path, phrase):
        pdf_file = open(file_path, 'rb')
        pdf_reader = PyPDF2.PdfFileReader(pdf_file)

        output = open(file_path.replace('pdf', 'csv'), 'w')
        output.write('SECTION, PARAGRAPH, LINE')

        current_section = None
        current_paragraph = None

        for page_number in range(pdf_reader.numPages):
            print('Current Page: %d' % page_number)
            page_text = pdf_reader.getPage(page_number).extractText()
            for line_number, line in enumerate(page_text.splitlines()):
                if 'SECTION' in line:
                    current_section = line.strip()
                    print('Current Section: %s' % current_section)
                elif len(line.strip()) >= 2 and line.strip()[1] == '.':
                    for character_number in range(len(line.strip())):
                        character = line.strip()[character_number]
                        if character_number % 2 != 0:
                            if character != '.':
                                break
                        else:
                            if not (character.isdigit() or character.isupper()):
                                break
                    else:
                        if line.strip()[-1] != '.':
                            current_paragraph = line.strip()
                            print('Current Paragraph: %s' % current_paragraph)
                elif phrase.lower() in line.lower():
                    if current_section is not None and current_paragraph is not None:
                        print('Phrase Found')
                        current_line = line.rstrip()
                        for next_line_number in range(line_number+1, line_number+10):
                            if next_line_number < len(page_text.splitlines()):
                                next_line = page_text.splitlines()[next_line_number]
                                if '.' in next_line:
                                    current_line += ' %s' % next_line[:(next_line.index('.') + 1)].strip()
                                    break
                                else:
                                    current_line += ' %s' % next_line.strip()
                        output.write('%s, %s, "%s"\n' % (current_section, current_paragraph, current_line))

        output.close() 
        popup = Popup(title='Completed',
                content=Label(text='The processing has been completed.\nResults are stored in:\n%s' % file_path.replace('pdf', 'csv')),
                size_hint=(None, None),
                size=(600, 200))
        popup.open()

class AutomatorApp(App):
    pass

if __name__ == '__main__':
    AutomatorApp().run()
