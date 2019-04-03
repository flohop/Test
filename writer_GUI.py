import csv
import os
import sys
import time
import winsound
import wikipedia
import wikipedia.exceptions
import wx
import wx.adv


file_dir = os.path.dirname(os.path.abspath(__file__))
base_path = os.path.dirname(file_dir) + "\WebScraping"
word_writer_file = base_path + "\word_writer"
email_sender_file = base_path + "\email_sender"

# my modules
from word_writer import WordWriter
import word_to_pdf
import email_sender, verifier
from email_sender import Verification
from wiki_api import WikiPage


dependencies = base_path + "\dependencies"

print(dependencies)


class StatusSetter():
    """To change the main status from another panel, change and acces a status here, to set later"""
    def set_status(self, status):
        self.global_status = str(status)

    def get_status(self):
        return self.gloabal_status


global_push_status = StatusSetter()


class WriterGui(wx.Frame):

    def __init__(self):
        wx.Frame.__init__(self, None, wx.ID_ANY, title="WikiLoader", size=(400, 400),
                          style=wx.DEFAULT_FRAME_STYLE ^ wx.RESIZE_BORDER)

        self.panel = wx.Panel(self, wx.ID_ANY)
        self.panel.SetSize(1000, 700)
        self.top_sizer = wx.BoxSizer(wx.VERTICAL)

        # Startbar and Statusbar

        icon_file = dependencies + "\mail_icon.png"
        icon1 = wx.Icon(icon_file, wx.BITMAP_TYPE_PNG)
        self.SetIcon(icon1)

        top_menu = wx.MenuBar()
        file_menu = wx.Menu()
        file_item = file_menu.Append(wx.ID_EXIT, 'Quit', 'Quit application')

        file_menu_2 = wx.Menu()
        file_item_2 = file_menu_2.Append(wx.ID_ANY, 'About me', 'Infos about the programmer')
        file_item_3 = file_menu_2.Append(wx.ID_ANY, "About the Software", "Infos about the application")

        top_menu.Append(file_menu, 'File')
        top_menu.Append(file_menu_2, "About")
        self.SetMenuBar(top_menu)

        self.Bind(wx.EVT_MENU, self.OnQuit, file_item)
        self.Bind(wx.EVT_MENU, self.OnAboutMe, file_item_2)
        self.Bind(wx.EVT_MENU, self.OnAboutSoftware, file_item_3)

        # Enter the wiki article name and shown if requested, a preview in a box

        self.text1 = wx.StaticText(self.panel, wx.ID_ANY, label="Gib den Artikel ein:")
        self.enter1 = wx.TextCtrl(self.panel, wx.ID_ANY)
        self.checkbox1 = wx.CheckBox(self.panel, wx.ID_ANY, label="Vorschau")

        self.Bind(wx.EVT_CHECKBOX, self.on_check, self.checkbox1)

        wiki_page_sizer = wx.BoxSizer(wx.HORIZONTAL)
        wiki_page_sizer.Add(self.text1, 1, wx.ALL, 13)
        wiki_page_sizer.Add(self.enter1, 1, wx.ALL, 10)
        wiki_page_sizer.Add(self.checkbox1, 1, wx.ALL, 13)

        # Box to show the preview of the article(To do: add loading animation)

        self.preview_box = wx.TextCtrl(self.panel, size=(400, 100),
                                       style=wx.TE_READONLY | wx.TE_MULTILINE | wx.ALIGN_CENTER)

        preview_box_sizer = wx.BoxSizer(wx.HORIZONTAL)
        preview_box_sizer.Add(self.preview_box, 1, wx.EXPAND | wx.ALL, 5)

        # Enter your address and get rest with a drobbox menu
        email_options = ["@gmail.com", "@yahoomail.com", "@gmxmail.com"]
        self.text2 = wx.StaticText(self.panel, wx.ID_ANY, label="Gib deine mail Addresse ein")
        self.enter2 = wx.TextCtrl(self.panel, wx.ID_ANY)
        self.combobox1 = wx.ComboBox(self.panel, wx.ID_ANY, choices=email_options, style=wx.CB_READONLY, value="@")

        mail_sizer = wx.BoxSizer(wx.HORIZONTAL)
        mail_sizer.Add(self.text2, 1, wx.ALL, 7)
        mail_sizer.Add(self.enter2, 1, wx.ALL, 5)
        mail_sizer.Add(self.combobox1, 1, wx.ALL, 5)

        # Button to send the email with the content
        self.send_button = wx.Button(self.panel, wx.ID_ANY, "Email senden")

        self.Bind(wx.EVT_BUTTON, self.on_button, self.send_button)

        button_sizer = wx.BoxSizer(wx.HORIZONTAL)
        button_sizer.Add(self.send_button, 1)

        # Clear Button
        self.checkbox2 = wx.CheckBox(self.panel, wx.ID_ANY, "Email merken")
        self.clear_button = wx.Button(self.panel, wx.ID_ANY, "Clear")

        self.Bind(wx.EVT_BUTTON, self.OnClear, self.clear_button)

        clear_sizer = wx.BoxSizer(wx.HORIZONTAL)
        clear_sizer.Add(self.clear_button, 1, wx.LEFT, 145)
        clear_sizer.Add(self.checkbox2, 1, wx.TOP | wx.LEFT, 5)

        # Status Bar
        self.status_bar = self.CreateStatusBar(2)
        self.status_bar.SetStatusWidths([-1, -4])
        self.status_bar.SetStatusText("Status:", 0)
        self.status_bar.SetStatusText("", 1)

        # Add all sizer to top sizer and set the top sizer
        self.top_sizer.Add(wiki_page_sizer)
        self.top_sizer.Add(preview_box_sizer)
        self.top_sizer.Add(mail_sizer)
        self.top_sizer.Add(button_sizer, 0, wx.CENTER | wx.TOP, 30)
        self.top_sizer.Add(clear_sizer, 0, wx.ALIGN_CENTER | wx.TOP, 30)

        self.SetSizer(self.top_sizer)

    def get_user_mail(self):
        return self.enter2.GetValue()

    def on_music(self, event):
        winsound.PlaySound(dependencies + "\sound_effect_1.wav", winsound.SND_ASYNC)

    def on_check(self, event):

        state = self.checkbox1.GetValue()
        if state:
            self.status_bar.PushStatusText(" Lade Vorschau...", 1)
            if self.enter1.GetValue() == "":
                self.checkbox1.SetValue(False)
                self.status_bar.PushStatusText("Bitte gebe zuerst einen Artikel ein", 1)

            else:
                # if box got checked and article name is given
                try:
                    article = WikiPage(self.enter1.GetValue())
                    article_summary = article.get_summary()
                    self.preview_box.SetValue(article_summary)
                    self.status_bar.PushStatusText("Vorschau geladen", 1)

                except:
                    self.preview_box.SetValue("Dieser Artikel konnte leider nicht gefunden werden.\n"
                                              "Bitte versuchen sie es mit einem anderen Artikel erneut.")
                    self.enter1.Clear()
        else:
            pass

    def on_button(self, event):
        """Create the .pdf and send it via email, if the mail is verrified """
        if not self.enter1.GetValue() == "" and not self.enter2.GetValue() == "" and not self.combobox1.GetValue() == "":
            msg = "Der Artikel '" + str(self.enter1.GetValue()) + "' wurde an " + str(self.enter2.GetValue()
                                                                                      + str(
                self.combobox1.GetValue()) + " gesendet.")

            # check if the email is in the file

            with open(dependencies + "\mail_verification.csv", "r") as ver_file:

                ver_reader = csv.reader(ver_file, delimiter=",")
                verified = False
                user_mail = self.enter2.GetValue() + self.combobox1.GetValue()
                for row in ver_reader:
                    try:
                        if row[0] == user_mail:
                            verified = True
                            break
                        else:
                            pass
                    except IndexError:
                        pass

            if not verified:

                self.status_bar.PushStatusText("Bitte verifizieren sie ihre Email:", 1)
                app_me = wx.App(False)
                frame_verification = VerifyEmailGUI(self.enter2.GetValue() + self.combobox1.GetValue()).Show()
                app_me.MainLoop()

            else:

                self.status_bar.PushStatusText("An die Arbeit...", 1)
                # mache all die Arbeit
                try:
                    page = WordWriter(self.enter1.GetValue(), "de")
                    self.status_bar.PushStatusText("Erstelle Word-Datei...", 1)
                except wikipedia.exceptions.DisambiguationError:
                    self.status_bar.PushStatusText("Fatal Error", 1)
                    sys.exit()
                except wikipedia.WikipediaException:
                    self.status_bar.PushStatusText("Bitte suche nach einen anderen Artikel")
                    self.enter1.Clear()


                try:
                    self.status_bar.PushStatusText("Erstelle Word-Datei...", 1)

                    # cuts out these sections from the article
                    slim_content = page.edit_content(page.wiki_page.get_content(page.wiki_page.get_summary()),
                                                     "Weblinks",
                                                     "Literatur", "Einzelnachweise", "Sonstiges", "Filme",
                                                     "Auszeichnungen", "Filmdokumentationen", "Anmerkungen",
                                                     "Biografien", "Weitere Texte")

                    to_mail = self.enter2.GetValue() + self.combobox1.GetValue()

                    page.add_title()
                    page.add_logo()
                    page.add_summary()
                    page.add_content(slim_content)
                    page.document.save(page.get_doc_name() + ".docx")
                    self.status_bar.PushStatusText("Konvertiere zu PDF...", 1)
                    page.delete_pic_in_outer_folder()
                    word_to_pdf.docx_to_pdf(page.get_doc_name() + ".docx", page.get_doc_name())
                    page.move_word_files()
                    self.status_bar.PushStatusText("Sende Email...", 1)
                    email_sender.send_email(to_mail, page.get_doc_name() + ".pdf", str(page.wiki_page.get_title()))
                    word_to_pdf.move_pdf_to_folder()
                    winsound.PlaySound(dependencies + "\sound_effect_1.wav", winsound.SND_ASYNC)
                    self.status_bar.PushStatusText("Email wurde erfolgreich gesendet!", 1)
                except PermissionError:
                    self.status_bar.PushStatusText("ERROR: Bitte schließe Word und versuche es erneut!", 1)
                except UnboundLocalError:
                    self.status_bar.PushStatusText("ERROR: Bitte wähle einen anderen Artikel aus", 1)

        else:
            msg = "Bitte füllen sie alle Felder aus"
            self.status_bar.PushStatusText(msg, 1)

    def OnQuit(self, event):
        self.status_bar.PushStatusText("Bye", 1)
        self.Close()

    def OnAboutMe(self, event):
        self.status_bar.PushStatusText("About me:", 1)
        app_me = wx.App(False)
        frame_me = AboutMeGUI().Show()
        app_me.MainLoop()

    def OnAboutSoftware(self, event):
        self.status_bar.PushStatusText("About the Software", 1)
        app_software = wx.App(False)
        frame_software = AboutSoftwareGUI().Show()
        app_software.MainLoop()

    def OnClear(self, event):
        self.enter1.SetValue("")
        self.preview_box.SetValue("")
        if self.checkbox2.GetValue():
            pass
        else:
            self.enter2.SetValue("")

        self.status_bar.PushStatusText("Cleared", 1)


class AboutMeGUI(wx.Frame):

    def __init__(self):
        wx.Frame.__init__(self, None, wx.ID_ANY, title="About me", size=(400, 400),
                          style=wx.DEFAULT_FRAME_STYLE ^ wx.RESIZE_BORDER)

        self.panel_me = wx.Panel(self, wx.ID_ANY)
        self.panel_me.SetSize(400, 400)

        icon_file_3 = dependencies + '\me_about_icon.png'
        icon3 = wx.Icon(icon_file_3, wx.BITMAP_TYPE_PNG)
        self.SetIcon(icon3)

        top_sizer = wx.BoxSizer(wx.VERTICAL)

        # Picture one
        try:
            jpg_img = dependencies + '\gates_bill_img.png'
            bmp1 = wx.Image(jpg_img, wx.BITMAP_TYPE_ANY).ConvertToBitmap()
            self.bitmap1 = wx.StaticBitmap(self.panel_me, -1, bmp1, (0, 0))  # Change coordinates later
        except IOError:
            print("Image not found")
            raise SystemExit

        # Line 1
        about_me = wx.StaticText(self.bitmap1, -1, "My Name is Florian Hoppe")
        font = wx.Font(14, wx.MODERN, wx.NORMAL, wx.BOLD,
                       underline=True)  # wx.Font(pointSize, family, style, weight, underline=False, faceName="", encoding=wx.FONTENCODING_DEFAULT)
        about_me.SetFont(font)

        line_1 = wx.BoxSizer(wx.HORIZONTAL)
        line_1.Add(about_me, 1, wx.TOP, 20)

        # Line 2
        about_me_2 = wx.StaticText(self.bitmap1, -1, "They call me the 'Bill Gates of Steinheim'")
        font_2 = wx.Font(10, wx.SWISS, wx.ITALIC, wx.NORMAL)
        about_me_2.SetFont(font_2)

        line_2 = wx.BoxSizer(wx.HORIZONTAL)
        line_2.Add(about_me_2)

        top_sizer.Add(line_1, 1, wx.ALIGN_CENTER, wx.TOP, 20)
        top_sizer.Add(line_2, 1, wx.ALIGN_CENTER)

        self.SetSizer(top_sizer)


class AboutSoftwareGUI(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None, wx.ID_ANY, title="About the Software", size=(400, 400),
                          style=wx.DEFAULT_FRAME_STYLE ^ wx.RESIZE_BORDER)

        icon_file_2 = dependencies + '\software_about_icon.png'
        icon2 = wx.Icon(icon_file_2, wx.BITMAP_TYPE_PNG)
        self.SetIcon(icon2)
        self.panel_me = wx.Panel(self, wx.ID_ANY)
        self.panel_me.SetSize(400, 400)

        top_sizer = wx.BoxSizer(wx.VERTICAL)

        try:
            jpg_image_2 = dependencies + '\software_img.png'
            bmp2 = wx.Image(jpg_image_2, wx.BITMAP_TYPE_ANY).ConvertToBitmap()
            self.bitmap_2 = wx.StaticBitmap(self.panel_me, -1, bmp2, (0, 0))
        except IOError:
            print("Image not found")
            raise SystemExit

        text1 = wx.StaticText(self.bitmap_2, -1, "Codename: ")
        font_1 = wx.Font(12, wx.SWISS, wx.NORMAL, wx.NORMAL, underline=True)
        text1.SetFont(font_1)

        sizer_line1 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_line1.Add(text1, 1, wx.TOP, 20)

        text2 = wx.StaticText(self.bitmap_2, -1, "ROCKET")
        text2.SetBackgroundColour((51, 153, 251))
        font_2 = wx.Font(50, wx.MODERN, wx.SLANT, wx.BOLD)
        text2.SetFont(font_2)

        sizer_line2 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_line2.Add(text2, 1)

        top_sizer.Add(sizer_line1, 1, wx.ALIGN_CENTER)
        top_sizer.Add(sizer_line2, 1, wx.ALIGN_CENTER)
        # ctr + v

        self.SetSizer(top_sizer)



class VerifyEmailGUI(wx.Frame):
    """Verifiziere die Email, falls sie noch nicht verifiziert ist"""

    def __init__(self, given_user_mail):
        wx.Frame.__init__(self, None, wx.ID_ANY, title="Verifikation", size=(400, 400),
                          style=wx.DEFAULT_FRAME_STYLE ^ wx.RESIZE_BORDER)

        self.verified = False
        self.panel_ver = wx.Panel(self, wx.ID_ANY)
        self.panel_ver.SetSize(400, 400)

        #Statusbar
        self.status_bar_2 = self.CreateStatusBar(2)
        self.status_bar_2.SetStatusWidths([-1, -4])
        self.status_bar_2.SetStatusText("Status:", 0)
        self.status_bar_2.SetStatusText("", 1)

        #Icon
        icon_file =  dependencies + '\icon_verify.png'
        icon = wx.Icon(icon_file, wx.BITMAP_TYPE_PNG)
        self.SetIcon(icon)

        # Send verification code and get the user_mail and the verif_code
        self.user_mail = given_user_mail
        send_ver = Verification(self.user_mail)

        send_ver.send_verify_mail()

        self.verif_code = send_ver.get_verif_code()

        # Create the GUI

        self.top_sizer = wx.BoxSizer(wx.VERTICAL)

        self.text1 = wx.StaticText(self.panel_ver, -1, "Es wurde ihnen eine Email mit dem Verifizierungscode "
                                                       "zugeschickt.")
        self.text2 = wx.StaticText(self.panel_ver, -1, "Bitte geben sie ihren Code unten ein.")

        text_sizer = wx.BoxSizer(wx.HORIZONTAL)
        text_sizer.Add(self.text1, 1, wx.TOP, 5)

        text_sizer_2 = wx.BoxSizer(wx.HORIZONTAL)
        text_sizer_2.Add(self.text2)

        self.enter_ver_code = wx.TextCtrl(self.panel_ver, -1)
        self.enter_ver_code.SetMaxLength(4)

        code_sizer = wx.BoxSizer(wx.HORIZONTAL)
        code_sizer.Add(self.enter_ver_code, wx.LEFT, 100)

        button_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.verify_button = wx.Button(self.panel_ver, -1, "Verifizieren")
        self.Bind(wx.EVT_BUTTON, self.on_check_button, self.verify_button)
        button_sizer.Add(self.verify_button, 1, wx.ALIGN_BOTTOM)

        self.exit_button = wx.Button(self.panel_ver, -1, "Abbrechen")
        self.Bind(wx.EVT_BUTTON, self.on_exit_button, self.exit_button)
        button_sizer.Add(self.exit_button, 1, wx.ALIGN_BOTTOM | wx.LEFT, 5)

        self.top_sizer.Add(text_sizer, 1, wx.ALIGN_CENTER)
        self.top_sizer.Add(text_sizer_2, 1, wx.ALIGN_CENTER | wx.TOP, 5)
        self.top_sizer.Add(code_sizer, 1, wx.ALIGN_CENTER)
        self.top_sizer.Add(button_sizer, 1, wx.ALIGN_RIGHT, wx.RIGHT, 5)
        self.SetSizer(self.top_sizer)

    def on_exit_button(self, event):
        self.Close()

    def on_check_button(self, event):
        if str(self.enter_ver_code.GetValue()) == str(self.verif_code):
            with open(dependencies + "\mail_verification.csv", 'a') as ver_file:
                ver_writer = csv.writer(ver_file, delimiter=',')
                ver_writer.writerow([self.user_mail, self.verif_code, "Verified"])
                self.status_bar_2.PushStatusText("Code korrekt, Email verifiziert!", 1)
                time.sleep(1)
                self.verified = True
                self.Close()

        else:
            #  Falls der Code nicht korrekt ist Status bar: Code nicht korrekt
            self.enter_ver_code.Clear()
            self.status_bar_2.PushStatusText("Code nicht korrekt", 1)

    def get_status(self):
        return self.verified


if __name__ == '__main__':
    app = wx.App(False)
    frame = WriterGui().Show()
    app.MainLoop()
