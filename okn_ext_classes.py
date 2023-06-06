# -*- encoding: utf-8 -*-#
# !/usr/bin/python

# Standardbibliotek import
import subprocess
import time

# Tredjeparts bibliotek import
import keyboard
import os
import pendulum
import pyperclip
import pypyodbc
from pywinauto.application import Application as PWA_App
import selenium.webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import WebDriverWait
import win32print

# Lokal applikasjon import
from okn_basic_classes import BasicWndHandler, MenuMaker, NamedList     # noqa
from okn_constants import MAMUT_RE, MOUTHPIECES_PROD_NUMS, PROGRAM_PATH, \
    SN_PROD_NUMS
from okn_constants import LIC_RENEWAL_PROD_NUMS
import okn_functions as okn


__author__ = 'Øyvind Nystad'

"""
Egenlagde klasser for ofte brukt funksjonalitet i Windows:
- MenuMaker
- WinGUIManager
- MamutManager
"""


class WinGUIManager(BasicWndHandler):
    """Klasse for å håndtere vinduer i Windows."""

    def archive_pdf_document(self,
                             folder_path,
                             file_name = None,
                             ) -> None:

        """
        Arkiver PDF-dokument i definert katalog.

        Hvis file_name er udefinert, settes filnavn lik som
        PDF-vindustittelen.
        """

        DOC_RE = '.*(?i)PDF.*'

        self.wnd_focus(title_re=DOC_RE, timeout=10.)


        document_wnd = (
            PWA_App(backend="uia").connect(title_re=DOC_RE,
                                           visible_only=True,
                                           found_index=0)
                                  .window(title_re=DOC_RE,
                                          visible_only=True,
                                          found_index=0)
                    )

        document_wnd.type_keys('^+s'
                               '{ENTER}',   # Velg annen mappe... OK
                               pause=0.4)
        self.await_text()

        # Gi opp dersom await_text() ikke oppdaterte utklippstavle
        if not pyperclip.paste():
            return None

        if not file_name:
            file_name = pyperclip.paste()

        pyperclip.copy(f'{folder_path}\\{file_name}')

        self.schedule_input_events(
            (0.4, 'ctrl+v'),
            (0.4, 'enter')
        )

        return None


    def await_text(self, *,
                   timeout: float=20.,
                   filltext: str='',
                   ) -> str:
        """Vent på fokus for aktivert tekst, returner så denne.

        Etter aktivering kan fokus på tekst (i tekstboks, e-postfelt
        etc.) ta noe tid. Forsøker kopiering av innhold til
        utklippstavle inntil dette lykkes, returnerer så tekstinnholdet.

        Argumenter:
            timeout: Antall sekunder før programmet gir opp å
                vente på fokus i tekstboks.
            filltext: Tekst forsøkt fylt inn i fokusert element.
        Retur:
            Innhold fra fokusert boks (str) dersom identifisert før
            timeout, ellers None.
        """

        T_PAUSE = 0.3   # For lav verdi gir ustabil oppdatering i Mamut

        def pyperclip_decorator(pyperclip_copy_func):
            """
            Dekoratør for pyperclip.copy.

            Sikrer at oppdatering av utklippstavle gjøres korrekt
            ved bruk av pyperclip.copy, da oppdatering ellers
            kan bruke noe tid, og ikke være fullført innen neste
            Python-kommando utføres.
            """
            def wrapper_func(text: str):
                """wrapper-funksjon."""
                text = str(text)
                pyperclip_copy_func(text)
                while pyperclip.paste() != text:
                    time.sleep(T_PAUSE)
                assert text == pyperclip.paste(), \
                    ("Utklippstavle ikke oppdatert, øk verdi for T_PAUSE, "
                     "evt. restart Mamut som kan ha stått på for lenge "
                     "og da begynt å bli lite responsivt.",
                     f"{pyperclip.paste()} {text}")
            return wrapper_func

        pyperclip_copy_new = pyperclip_decorator(pyperclip.copy)

        t_start = time.time()
        if filltext:
            pyperclip_copy_new('')
            while not pyperclip.paste():
                # Kanskje nødvendig, reduser evt. senere
                # time.sleep(0.25)
                keyboard.press_and_release('0')
                # Kanskje nødvendig, reduser evt. senere
                # time.sleep(1.25)
                time.sleep(T_PAUSE)
                keyboard.press_and_release('ctrl+a+c')
                time.sleep(T_PAUSE)
                assert time.time() - t_start < timeout, \
                    f"Operasjonen tok over {timeout} sek - avslutter."
            time.sleep(T_PAUSE)

            keyboard.write(str(filltext))

            time.sleep(T_PAUSE)
            return filltext
        else:
            pyperclip_copy_new('')
            while not pyperclip.paste():
                keyboard.press_and_release('ctrl+a+c')
                time.sleep(T_PAUSE)        # Nødvendig pause
                if time.time() - t_start > timeout:
                    input(f"Operasjonen tok mer enn {timeout} sek - gir opp\n"
                          "Gi gjerne beskjed til utvikler om problemet,\n"
                          "Trykk ENTER for å forsøke å fullføre.")
                    return ''

            return pyperclip.paste().strip()

    def compose_outlook_email(self, *,
                              email: str ='',
                              cc: str='',
                              subject: str='',
                              body: str='',
                              attach: str='') -> None:

        """Opprett e-post i Outlook."""

        print(f"\nOppretter e-post med emne '{subject}'... ", end="")

        # Windows hex-verdi som tilsvarer Python linefeed \n: 0D 0A
        body.replace('\n', '%0D%0A')

        if not self.wnd_focus(title_re=r'Outlook.*', is_maximized=True):
            print('Outlook må være åpen. Avslutter...'.ljust(44))
            time.sleep(3.)
            return None

        cc_arg = f'cc={cc}' if cc else None
        subject_arg = f'subject={subject}' if subject else None
        body_arg = f'body={body}' if body else None

        email_args = '&'.join(filter(None, [cc_arg, subject_arg, body_arg]))

        okn.start_winprog(
            name='outlook',
            param=(f'/c ipm.note /m "mailto:{email}?{email_args}"' +
                   (f' /a "{attach}"' if attach else ''))
            )

        self.wnd_focus(title_re=f'{subject} - Melding (HTML).*',
                       is_maximized=True)
        print("OK")
        return None


    def get_web_control(self, browser, x_path=''):
        """
        Finn web-element basert på XPath, avvent at dette blir klikkbart
        og dermed klar for handling, og returner web-elementet

        'browser' er element av type selenium.webdriver.Firefox

        Returnerer None hvis klikkbart element ikke ble funnet i tide

        Angående XPath-verdier:
        Dette er koder fra websiden som entydig identifiserer
        hvert element. Man kan finne ut elementets Xpath i Firefox
        ved: Høyreklikk -> Undersøk -> Klikk
        på valgt element -> Høyreklikk -> Copy -> XPath.
        """

        MAX_WAIT_TIME = 5
        browser_wait = WebDriverWait(browser, MAX_WAIT_TIME)

        try:
            browser_wait.until(
                expected_conditions.element_to_be_clickable(
                    (By.XPATH, x_path)
                    )
                )
            time.sleep(0.2)                 # Kanskje nødvendig pause
            return browser.find_element_by_xpath(x_path)
        except Exception:
            return None


    def get_printer_names(self) -> list:
        """Returner liste over tilgjengelige printernavn."""
        printer_info = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)
        printer_names = [name for (flags, description, name, comment) in
                         printer_info]
        return printer_names


    def schedule_input_events(self,
                              *args) -> None:
        """
        Utfør hendelser på brukergrensesnittet, f.eks. knappetrykk
        """
        for arg in args:
            assert isinstance(arg, tuple), \
                (f"Forventet argument av type tuple, mottok "
                 f"{type(arg).__name__}")
            repetition_count = arg[2] if len(arg) >= 3 else 1
            pressed_key = arg[1] if len(arg) >= 2 else None
            for __ in range(repetition_count):
                time.sleep(arg[0])
                if pressed_key:
                    assert isinstance(pressed_key, str), \
                        ("pressed_key er av type "
                         f"'{type(pressed_key).__name__}', "
                         "forventet 'str'")
                    keyboard.press_and_release(pressed_key)
        return None


    def set_webdriver(self) -> selenium.webdriver.Firefox:
        """
        Initialiser Geckodriver, webdriver for å kontrollere
        Firefox-nettleser
        """
        webdriver = None
        try:
            print("\nStarter Firefox via Geckodriver... ", end="")
            os.chdir(fr'{PROGRAM_PATH}\Resources')
            # Denne linjen feiler dersom browser gjør en oppdatering
            # ved oppstart
            webdriver = selenium.webdriver.Firefox(
                executable_path=r'geckodriver.exe')
            print("OK")                         # Startet Firefox
        except selenium.common.exceptions.WebDriverException:
            input("\nProblem med initialisering av webdriver.\n"
                  "Vanligste årsak er at Firefox var opptatt med å\n"
                  "oppdatere til ny versjon. Steng Firefox, trykk\n"
                  "ENTER for å gå tilbake til hovedmeny, og forsøk\n"
                  "en gang til")
        return webdriver


class MamutManager:
    """Klasse for å benytte funksjonalitet i Mamut."""

    def __init__(self):
        self.wingui = WinGUIManager()
        self.app_gui = None
        self.curr_order_num = None
        self.curr_order_is_invoiced: bool = False
        self.is_alive: bool = False

    def lookup_db(self, sql_statement):
        """Eksekver SQL-spørring mot Mamut-database.

        Returnerer liste av NamedList (en record per element, nøkkel
        er attributtnavn. Dersom ingen record, returneres None.
        """
        def autoconvert(val):
            """Typecaster verdi til mest passende type.

            Verdier hentet via SQL-spørring har allerede variabeltypen
            slik den er definert i Mamut-databasen, men det kan være
            hensiktsmessig å overstyre typen og bearbeide/parse
            verdiene ytterligere. Returnerer verdi av type
            bool, int, float, pendulum.datetime eller str.

            Eksempel:
                1234 -> int
                '1234' -> int
                '1234.0' -> int
                '1234.5' -> float
                '01234' -> str
                '0' -> int
            """
            if isinstance(val, bool):
                assert type(val).__name__ == 'bool', "Feil type"
            elif val is None:
                assert type(val).__name__ == 'NoneType', "Feil type"
            else:
                # Uvisst om replace er nødvendig
                val = str(val).strip().replace('\r', '\n')

                if all([val != '0', val.startswith('0'),
                        not val.startswith('0.')]):
                    assert type(val).__name__ == 'str', "Feil type"
                else:
                    try:
                        val = float(val)
                    except ValueError:      # Dersom str (ikke int/float)
                        try:
                            val = pendulum.from_format(
                                val, 'YYYY-MM-DD HH:mm:SS')
                            assert type(val).__name__ == 'DateTime', \
                                "Feil type"
                        except ValueError:
                            assert type(val).__name__ == 'str', "Feil type"
                    else:                   # Dersom tall (int/float)
                        if val.is_integer():
                            val = int(val)
                            assert type(val).__name__ == 'int', "Feil type"
                        else:
                            assert type(val).__name__ == 'float', "Feil type"
            return val

        conn = pypyodbc.connect(r'Driver={SQL Server};'
                                r'Server=[SERVER NAME]t;'
                                r'Database=[DB NAME];'
                                r'uid=[UID];'
                                r'pwd=[PWD]')
        cursor = conn.cursor()
        cursor.execute(sql_statement)
        query_raw = list(cursor.fetchall())

        # Liste over kolonnenavn
        col_headers = [var_tuple[0] for var_tuple in cursor.description]

        cursor.close()
        del cursor
        conn.close()

        # Lag liste av NamedList. Ett listeelement = 1 db-record.
        # Attributt aksesseres med syntaks
        # [namedlist_navn].[attributt_navn]
        query_refined = []
        for record in (query_raw):
            attribs = NamedList()
            for idx, elem in enumerate(record):
                setattr(attribs, col_headers[idx], autoconvert(elem))
            query_refined.append(attribs)
        if len(query_refined) == 0:
            return [None]
        else:
            return query_refined

    def scan_order_num(self):
        """
        Henter ordrenummer i åpen Mamut-ordre via pywinauto-objekt.
        """
        if not self.is_alive:
            return None
        else:
            try:
                print("Detekterer Mamut-ordrenummer... ", end="")
                txt_short_name = (self.gui_app
                                      .child_window(title='txtShortName',
                                                    control_type='Edit')
                                  )
                order_or_invoice_str = txt_short_name.iface_value.CurrentValue
                txt_short_name.draw_outline(colour='blue', thickness=3)
                time.sleep(1.)
            except Exception:
                print("feilet")
                order_or_invoice_str = ''

        if any([s in order_or_invoice_str for s in [
                "Annullert",
                "Kreditordre",
                "Ordre",
                "Restordre",
                "Samleordre",
                ]
               ]):

            self.curr_order_num = int(order_or_invoice_str.split()[1])
            print(self.curr_order_num)
            self.curr_order_is_invoiced = False

        elif "Faktura" in order_or_invoice_str:
            invoice_number = order_or_invoice_str.split()[1]
            self.curr_order_is_invoiced = True
            self.curr_order_num: int = self.lookup_db(f"""
                 SELECT orderid FROM g_order
                 WHERE invoiceid = {invoice_number}
            """)[0].orderid
            print(self.curr_order_num)
            print(f"Detektert fakturanummer: {invoice_number}")
        else:
            self.curr_order_num = None
            print()
            okn.mention_return_to_main_menu(
                "Mamut er aktiv, men ikke i åpen ordre")
        return None

    def update_sys_info(self):
        """Oppdaterer Mamut-systeminfo

        Oppdaterer:
        - self.is_alive: True/False, om Mamut er åpen
        - self.gui_app: pywinauto-objekt for å kontrollere dialoger i Mamut
        """
        print("Kobler til Mamut-applikasjon... ", end="")
        if 'Mamut.exe' not in subprocess.getoutput('tasklist'):
            print("feilet.\n"
                  "Mamut.exe må være aktiv.")
            okn.mention_return_to_main_menu()
            self.is_alive = False
            self.gui_app = None
        else:
            print("OK")
            self.wingui.wnd_focus(title_re=MAMUT_RE,
                                  is_maximized=True)['wnd_handle']
            self.is_alive = True
            self.gui_app = (PWA_App(backend="uia").connect(title_re=MAMUT_RE,
                                                           found_index=0)
                                                  .window(title_re=MAMUT_RE,
                                                          found_index=0)
                            )
        return None

    def open_customer(self, *,
                      cust_num,
                      action=None,
                      ):
        """
        Henter opp ønsket kunde i Mamut.
        Mulige parametere:
            action: Angir om det opprettes ny ordre, eller benyttes
                eksisterende ordre
        """
        clipboard = pyperclip.paste().strip()   # Spar utklippstavle

        allowed_actions = ['create_new', 'open_existing', None]
        assert action in allowed_actions, \
            f"Forventet verdi fra {allowed_actions} mottok {action}"

        # assert self.is_alive, "Mamut er ikke aktiv"
        if not self.is_alive:
            self.update_sys_info()
            assert self.is_alive, "Mamut er ikke aktiv"

        self.wingui.wnd_focus(title_re=MAMUT_RE, is_maximized=True)

        customer_name = self.lookup_db(f"""
             SELECT name FROM g_contac
             WHERE custid = {cust_num}
        """)[0].name
        pyperclip.copy(customer_name)

        print(f"Henter kunde {customer_name}... ", end="")

        self.gui_app.type_keys(
            '^s'                    # ctrl+s -> Lagre
            '%i'                    # alt+i -> Vis
            'k'                     # alt+k -> Kontakt
            'k'                     # alt+k -> Kontaktoppfølging
            '^s'                    # ctrl+s -> Lagre
            '^l'                    # ctrl+l -> Liste
            '^v',                   # ctrl+v -> Lim inn customer_name
            pause=0.3)              # Øk pause dersom problem

        self.gui_app.child_window(title='OK').type_keys('{ENTER}')

        # Fokus på Ordre/Faktura-tabkort
        (self.gui_app
             .child_window(title='Kontaktoppfølging')
             .child_window(title='PageFrame', control_type='Tab')
             .type_keys('{RIGHT 3}')
         )
        print("OK")                     # Ferdig hentet kunde

        if action == 'create_new':
            print("Oppretter ny Mamut-ordre... ", end="")
            (self.gui_app
                 .child_window(title='New',
                               control_type='Group',
                               found_index=0,
                               )
                 .click_input()
             )
            # Muligens nødvendig pause for å hindre at annen åpen ordre
            # hentes opp
            time.sleep(0.2)
            print("OK")                 # Ferdig opprettet Mamut-ordre
        elif action == 'open_existing':
            # Feltnavn kan variere, derfor (?i) = case-insensitive
            print("Åpner eksisterende Mamut-ordre... ", end="")
            (self.gui_app
                 .child_window(title_re='(?i)cmbOrderStatus',
                               control_type='ComboBox')
                 .type_keys('{SPACE 2}'     # Aktiver og ekspander komboboks
                            '{PGUP}'        # Fokus på øverste element
                            'u'             # Fokus på 'Ubehandlet ordre'
                            '{SPACE}'       # Velg
                            '^+r',          # ctrl+shift+r -> Rediger ordre
                            pause=0.25,
                            )
             )
            print("OK")                 # Ferdig åpnet eksisterende Mamut-ordre

        pyperclip.copy(clipboard)       # Tilbakefør oppr. utklippstavle
        time.sleep(0.2)                 # For sikkerhets skyld
        return None

    def get_ordered_prods(self):
        """
        Leser ordrelinjer fra åpen Mamut-ordre etter SQL-spørring
        mot database, og deler resultatet inn i grupper:
        - Munnstykker
        - Andre lagervarer
        - Ikke-lagervarer
        Mulig innparameter: Mamut-ordrenummer
        """

        assert self.curr_order_num, "Ikke gyldig ordrenummer"

        mamut_order_prods = NamedList(
            mouthpcs=dict(),
            sn_devices=dict(),
            other_stor=dict(),
            non_stor=dict(),
            has_FP00_prod=False,
            has_only_non_stor_prods=False,
            has_only_mouthpiece_prods=False,
            has_lic_renewal_prods=False,
            num=self.curr_order_num
        )

        prod_qry = self.lookup_db(f"""
            SELECT g_orderl.qtyorder,
            g_orderl.prodid AS prod_num,
            g_prod.usestore
            FROM g_orderl
            JOIN g_order ON g_orderl.linkid=g_order.linkid
            JOIN g_prod ON g_prod.prodid=g_orderl.prodid
            WHERE g_order.orderid = {mamut_order_prods.num}
            AND g_orderl.repstrucorder = 0  /* Neglisjér strukturvare-produkt*/
        """)

        if prod_qry != [None]:

            for elem in prod_qry:
                # Munnstykker
                if elem.prod_num in MOUTHPIECES_PROD_NUMS:
                    if elem.prod_num not in mamut_order_prods.mouthpcs.keys():
                        mamut_order_prods.mouthpcs[elem.prod_num] = 0.
                    mamut_order_prods.mouthpcs[elem.prod_num] += elem.qtyorder
                # Apparat med serienummer
                elif elem.prod_num in SN_PROD_NUMS:
                    if (elem.prod_num not in
                       mamut_order_prods.sn_devices.keys()):
                        mamut_order_prods.sn_devices[elem.prod_num] = 0.
                    # Bruk av abs() for ordrer der noteres f.eks. -1 apparater,
                    # for returer, slik at dette telles som 1
                    mamut_order_prods.sn_devices[elem.prod_num] += abs(
                        elem.qtyorder)
                # Øvrige lagervarer
                elif elem.usestore:  # Lagervare
                    if (elem.prod_num not in
                       mamut_order_prods.other_stor.keys()):
                        mamut_order_prods.other_stor[elem.prod_num] = 0.
                    mamut_order_prods.other_stor[elem.prod_num] += \
                        elem.qtyorder
                # Ikke-lagervarer
                else:
                    if elem.prod_num not in mamut_order_prods.non_stor.keys():
                        mamut_order_prods.non_stor[elem.prod_num] = 0.
                    if elem.prod_num == 'FP00':
                        mamut_order_prods.has_FP00_prod = True
                    if elem.prod_num in LIC_RENEWAL_PROD_NUMS:
                        mamut_order_prods.has_lic_renewal_prods = True
                    mamut_order_prods.non_stor[elem.prod_num] += elem.qtyorder

        mamut_order_prods.has_stor_prods = True if any(
            (mamut_order_prods.mouthpcs,
             mamut_order_prods.sn_devices,
             mamut_order_prods.other_stor,
             )
        ) else False

        if (mamut_order_prods.non_stor
           and not mamut_order_prods.mouthpcs
           and not mamut_order_prods.sn_devices
           and not mamut_order_prods.other_stor):
            mamut_order_prods.has_only_non_stor_prods = True

        mamut_order_prods.has_only_mouthpiece_prods = True if all(
            (mamut_order_prods.mouthpcs,
             not mamut_order_prods.sn_devices,
             not mamut_order_prods.other_stor,
             )
        ) else False

        print("\nVARER I MAMUT-ORDRE PER KATEGORI")

        print("Munnstykker:", end="")
        if mamut_order_prods.mouthpcs:
            for prod_num in mamut_order_prods.mouthpcs:
                print('\r\t\t\t'
                      f'{mamut_order_prods.mouthpcs[prod_num]} x {prod_num}')
        else:
            print('\r\t\t\t---')

        print("Apparater:", end="")
        if mamut_order_prods.sn_devices:
            for prod_num in mamut_order_prods.sn_devices:
                print('\r\t\t\t'
                      f'{mamut_order_prods.sn_devices[prod_num]} x {prod_num}')
        else:
            print('\r\t\t\t---')

        print("Andre lagervarer:", end="")
        if mamut_order_prods.other_stor:
            for prod_num in mamut_order_prods.other_stor:
                print('\r\t\t\t'
                      f'{mamut_order_prods.other_stor[prod_num]} x {prod_num}')
        else:
            print('\r\t\t\t---')

        print("Ikke-lagervarer:", end="")
        if mamut_order_prods.non_stor:
            for prod_num in mamut_order_prods.non_stor:
                print('\r\t\t\t'
                      f'{mamut_order_prods.non_stor[prod_num]} x {prod_num}')
        else:
            print('\r\t\t\t---')

        return mamut_order_prods

    def get_order_properties(self, order_num=None):
        assert order_num, "Ikke gyldig ordrenummer"

        # SQL-setningen kan gi flere treff, en for hvert postnummer hvis
        # kunden har registrert flere leveringsadresser. I praksis har
        # det ikke betydning for fraktberegning, da sonenummer (1-5)
        # uansett blir det samme

        order_properties = self.lookup_db(f"""
            SELECT
            g_clisys.descr                 AS lev_betingelser,
            g_contac.countrycodecustomer   AS country_id,   /* Norge = 1*/
            g_contac.email                 AS cust_email,
            g_contac.enterno               AS org_num,
            g_currency.isocode             AS currency,
            g_order.custid                 AS cust_num,
            g_order.contname               AS cust_name,
            g_order.curr_sum_n             AS brutto_sum,
            g_order.data67                 AS avrunding_id,
            g_order.datedeliv              AS lev_dato,
            g_order.dateinvoice            AS fakturadato,
            g_order.electronicdocumenttype AS is_ehf_invoice,
            g_order.freightvolumesum       AS volume,
            g_order.ifactoringstatus       AS is_factoring,
            g_order.invoiceid              AS invoice_num,
            g_order.lorderready            AS klar_til_fakturering,
            g_order.maincontid             AS main_office_contact_num,
            g_order.maincontname           AS main_office_name,
            g_order.maincontres            AS is_main_office_invoiced,
            g_order.refyour                AS deres_ref,
            g_order.reference              AS referanse,
            g_order.reportidinvoice        AS formular_id,
            g_deli.zipcode                 AS zip_code,
            w_delitypes.[freetext]         AS lev_form
            FROM g_order
            JOIN g_clisys    ON g_order.data7 = g_clisys.nr
            JOIN g_deli      ON g_deli.sourceid = g_order.contid
            JOIN g_contac    ON g_contac.custid = g_order.custid
            JOIN g_currency  ON g_order.currencyid = g_currency.currencyid
            JOIN w_delitypes ON w_delitypes.uniqueid = g_order.data2
            WHERE g_clisys.id = 7
            AND g_deli.adrtype = 1
            AND g_order.orderid = {order_num}
        """)[0]

        def _trim_postal_numbers(order_properties=order_properties):
            """
            Fjern eventuelle mellomrom i postnummer. Dette kan f.eks.
            forekomme på svenske postnumre.
            """
            order_properties.zip_code = str(
                order_properties.zip_code).replace(' ', '')
            return order_properties


        def _add_main_office_info(order_properties=order_properties):
            """
            Legg hovedkontor-info til ordreegenskaper
            """
            if order_properties.main_office_contact_num:
                order_properties = order_properties.update(self.lookup_db(
                    f"""
                    SELECT
                    g_contac.custid    AS main_office_cust_num,
                    g_contac.vend      AS has_vendor_main_office,
                    g_contac.cooporate AS has_dealer_main_office,
                    g_contac.enterno   AS main_office_org_num
                    FROM g_contac
                    WHERE g_contac.contid =
                        {order_properties.main_office_contact_num}
                """)[0]
                )
            else:
                order_properties = order_properties.update(
                    has_vendor_main_office = False,
                    has_dealer_main_office = False,
                    main_office_org_num = None,
                    main_office_name = None,
                )

            return order_properties

        def _add_contact_pers_info(order_properties=order_properties):
            """
            Legg kontaktperson-e-post til ordreegenskaper
            """
            order_properties.deres_ref_email = None
            if order_properties.deres_ref:
                order_properties = order_properties.update(self.lookup_db(
                    f"""
                    SELECT email AS deres_ref_email
                    FROM g_cpers WHERE
                    CONCAT(TRIM(FIRSTNAME), ' ', TRIM(LASTNAME)) =
                    '{order_properties.deres_ref}'
                    """)[0]
                )
            return order_properties


        order_properties = _trim_postal_numbers(order_properties)
        order_properties = _add_main_office_info(order_properties)
        order_properties = _add_contact_pers_info(order_properties)


        assert order_properties is not None, (
            "order_properties=None.\n"
            "Dette kan skje hvis Leveringsform=(Ingen).\n"
            "Endre denne hvis dette var tilfelle."
        )

        # For benevning i cm3, absoluttverdi fordi volum kan bli
        # negativt hvis antall produkter er negativt (aktuelt for
        # retur av apparater
        order_properties.volume = abs(order_properties.volume * 1000.)
        order_properties.volweight = order_properties.volume / 5.

        # g_order.ifactoringstatus er tallverdi 0/1 - ønsker True/False
        order_properties.is_factoring = bool(order_properties.is_factoring)

        # kolon i filnavn gir problemer
        order_properties.referanse = (
            str(order_properties.referanse).replace(':', ';')
        )

        order_properties.is_export_shipment = (
            True if order_properties.country_id != 1 else False)
        del order_properties.country_id

        order_properties.is_ehf_invoice = bool(order_properties.is_ehf_invoice)
        order_properties.formula = {
            4410: 'Faktura u/giro',
            4401: 'Internasjonal faktura',
        }.get(order_properties.formular_id)

        del order_properties.formular_id

        # TODO: Omgå at order_properties = None hvis Leveringsform = (Ingen)

        return order_properties


    def get_prod_num_by_sn(self, serial_num):
        """
        Finn Mamut produktnummer ut fra serienummer.
        """

        prod_num = self.lookup_db(f"""
            SELECT g_prod.prodid FROM g_prod
            WHERE g_prod.pk_prodid = (
                SELECT TOP 1 g_storeitem.fk_product FROM g_storeitem
                WHERE g_storeitem.serialnr = '{serial_num}'
                ORDER BY g_storeitem.fk_product DESC
                )
            """)[0].prodid
        return prod_num



    def _save_order(self):
        # Lagre Mamut-ordre
        print("Lagrer ordre... ", end="")
        self.wingui.schedule_input_events(
            (0.5, 'ctrl+s'),
            (1.0, None),
        )
        print("OK")                     # Ferdig lagret ordre
        return None

    def set_order_properties(self, *,
                             deres_ref=None,
                             lev_betingelser=None,
                             lev_dato=None,
                             lev_form=None,
                             formular=None,
                             referanse=None,
                             faktura_tekst=None,
                             pakkseddel_tekst=None,
                             tab=None,
                             use_default_misc_settings=False,
                             ) -> None:
        """
        Fyller inn ønskede verdier i Mamut-ordre, dersom
        dersom feltet ikke har ønsket verdi allerede.
        Mulige innparametere:
        - ordrenummer
        - 'deres_ref', 'lev_betingelser', 'lev_dato',
          'lev_form', 'formular', 'referanse',
          'faktura_tekst', 'pakkseddel_tekst',
        - tab. Spesifiserer tabkort som skal være aktivert
        ('Produktlinjer', 'Frakt', 'Tekst' eller 'Diverse')
        """
        if self.curr_order_num is None:
            self.scan_order_num()

        assert self.curr_order_num, \
            f"Ikke gyldig ordrenummer: {self.curr_order_num}"

        # Spar opprinnelig utklippstavle
        clipboard = pyperclip.paste().strip()

        allowed_tab_vals = ['Produktlinjer', 'Frakt', 'Tekst',
                            'Diverse', None]

        assert tab in allowed_tab_vals, (
            f"Forventet verdi blant {allowed_tab_vals}, mottok: {tab}")

        assert use_default_misc_settings in [False, True], (
            f"Forventet verdi lik False eller True, mottok: "
            f"{use_default_misc_settings}")

        order_properties = self.get_order_properties(self.curr_order_num)

        def _change_tab(chosen_tab=tab):
            """Endre fokusert tabkort."""
            assert chosen_tab is not None, \
                "chosen_tab kan ikke være None"
            key_presses_by_card_choice = {
                'Produktlinjer': '{LEFT}{RIGHT}',
                'Frakt': '{RIGHT}',
                'Tekst': '{RIGHT 2}',
                'Diverse': '{LEFT}'
                }

            (self.gui_app
                 .child_window(title='clsPageFrame')
                 .child_window(title='PageFrame')
                 .type_keys(key_presses_by_card_choice[chosen_tab],
                            pause=0.1)
             )
            return None

        # Setter innstillinger under "Frakt"
        if any(
            (lev_betingelser not in [None, order_properties.lev_betingelser],
             lev_form not in [None, order_properties.lev_form])
             ):
            _change_tab(chosen_tab='Frakt')
            if (lev_betingelser not in
                    [None, order_properties.lev_betingelser]):
                print(f"Endrer leveringsbetingelser... ", end="")
                deliv_cond_wnd = self.gui_app.child_window(title='cmbData7')
                deliv_cond_wnd.draw_outline(colour='blue', thickness=3)
                deliv_cond_wnd.click_input()
                time.sleep(0.1)
                deliv_cond_wnd.type_keys(lev_betingelser[:3])
                print("OK")             # Ferdig endret leveringsbetingelse
            if lev_form not in [None, order_properties.lev_form]:
                print(f"Endrer leveringsform... ", end="")
                deliv_type_wnd = self.gui_app.child_window(title='cmbData2')
                deliv_type_wnd.draw_outline(colour='blue', thickness=3)
                deliv_type_wnd.click_input()
                time.sleep(0.1)
                deliv_type_wnd.type_keys(lev_form[:3])
                print("OK")             # Ferdig endret leveringsform

        if use_default_misc_settings:
            # Sett standardinnstillinger under "Diverse"
            correct_formula = (
                'Internasjonal faktura' if order_properties.is_export_shipment
                else 'Faktura u/giro'
                )

            if (order_properties.currency == 'SEK' and
                order_properties.is_ehf_invoice):
                advised_factoring_setting = True
            else:
                advised_factoring_setting = False

            NO_ROUNDOFF_ID = 1
            if any((order_properties.avrunding_id != NO_ROUNDOFF_ID,
                    order_properties.formula != correct_formula,
                    not order_properties.klar_til_fakturering,
                    order_properties.is_factoring != advised_factoring_setting,
                    (order_properties.is_ehf_invoice and
                     order_properties.brutto_sum == 0),
                    )):
                _change_tab(chosen_tab='Diverse')
                print()
                if order_properties.avrunding_id != NO_ROUNDOFF_ID:
                    # input(order_properties.avrunding_id)
                    # TAB-trykk nødvendig for å omgå bug i Mamut hvor
                    # ordre spontant får status 'Annullert'
                    print("Fjerner avrunding-innstilling... ", end="")
                    roundoff_wnd = self.gui_app.child_window(title='cmbData67')
                    roundoff_wnd.draw_outline(colour='blue', thickness=3)
                    roundoff_wnd.type_keys('{PGUP}'     # Velg (Ingen)
                                           '{TAB}', pause=0.1)
                    print("OK")         # Ferdig med fjerning av avrunding

                if order_properties.formula != correct_formula:
                    # Velg element over 'Internasjonal faktura'
                    print("Korrigerer fakturaformular... ", end="")
                    inv_form_wnd = self.gui_app.child_window(title='cmbReport')
                    inv_form_wnd.draw_outline(colour='blue', thickness=3)
                    inv_form_wnd.type_keys('{SPACE}'
                                           'i' +
                                           '{UP}' * (correct_formula ==
                                                     'Faktura u/giro') +
                                           '{TAB}', pause=0.1)
                    print("OK")         # Ferdig endret fakturaformular

                if not order_properties.klar_til_fakturering:
                    # click_input() heller enn type_keys('{SPACE}') da
                    # element kan være utilgjengelig for tastetrykk
                    print("Krysser av for Klar til fakturering... ", end="")
                    inv_ready_wnd = (
                        self.gui_app
                            .child_window(title='Klar til fakturering')
                        )
                    inv_ready_wnd.draw_outline(colour='blue', thickness=3)
                    # Forsøk med følgende kommandoer var ustabile:
                    # inv_ready_wnd.click_input()
                    # inv_ready_wnd.type_keys('{SPACE}')
                    inv_ready_wnd.set_focus()
                    # Nødvendig pause for å omgå bug i Mamut (ordre skifter
                    # spontant status til annullert). Øk pause om nødvendig.
                    time.sleep(0.4)
                    keyboard.press_and_release('space')
                    print("OK")         # Ferdig avkrysset

                if order_properties.is_factoring != advised_factoring_setting:
                    factoring_wnd = self.gui_app.child_window(
                        title='cboFactoring')
                    factoring_wnd.draw_outline(colour='blue', thickness=3)
                    if advised_factoring_setting == True:
                        print("Slår på Factoring-innstilling... ", end="")
                        factoring_wnd.type_keys('j', pause=0.1)     # j for Ja
                    if advised_factoring_setting == False:
                        print("Slår av Factoring-innstilling... ", end="")
                        factoring_wnd.type_keys('n', pause=0.1)     # n for Nei
                    factoring_wnd.click_input()
                    print("OK")

                # Må slå av innstilling for EHF-faktura hvis ordresum er 0,
                # dvs. at det lages nullfaktura
                if (order_properties.is_ehf_invoice and
                    order_properties.brutto_sum == 0):
                    e_format_wnd = self.gui_app.child_window(
                        title='cboElectronicType')
                    e_format_wnd.draw_outline(colour='blue', thickness=3)
                    print("Slår av EHF-innstilling... ", end="")
                    e_format_wnd.type_keys('{PGUP}'     # Velg (Ingen)
                                           '{TAB}', pause=0.1)
                    print("OK")



        # Sett inn fast tekst (fakturatekst) og tekst på pakkseddel
        if any([faktura_tekst, pakkseddel_tekst]):
            _change_tab(chosen_tab='Tekst')
            if faktura_tekst:
                print("Fyller inn fakturatekst (Fast tekst)... ", end="")
                pyperclip.copy(faktura_tekst)
                (self.gui_app
                     .child_window(title='cmbTextType')
                     .type_keys('{PGUP}'))
                (self.gui_app
                     .child_window(title='txtEditMemo')
                     .type_keys('^v'))
                print("OK")             # Ferdig fylt inn fakturatekst
            if pakkseddel_tekst:
                print("Fyller inn pakkseddel-tekst... ", end="")
                pyperclip.copy(pakkseddel_tekst)
                (self.gui_app
                     .child_window(title='cmbTextType')
                     .type_keys('{PGUP}{DOWN 2}'))
                (self.gui_app
                     .child_window(title='txtEditMemo')
                     .type_keys('^v'))
                print("OK")             # Ferdig fylt inn pakkseddel-tekst

        if lev_dato:
            # Case-insensitivt søk på txtDate og txtYourRef fordi feltene kan
            # få navn som txtdate og txtYourref (bug i Mamut)
            print("Fyller inn Leveringsdato... ", end="")
            deliv_date_wnd = (
                self.gui_app
                    .child_window(title='txtDateDeliv')     # Nødvendig linje
                    .child_window(title_re='(?i)txtDate',
                                  control_type='Edit')
                )
            deliv_date_wnd.draw_outline(colour='blue', thickness=3)
            deliv_date_wnd.type_keys(lev_dato)
            print("OK")                 # Ferdig fylt inn leveringsdato

        if referanse:
            print("Fyller inn Referanse... ", end="")
            ref_wnd = (
                self.gui_app.child_window(title='txtReferance',
                                          control_type='Edit')
                )
            ref_wnd.draw_outline(colour='blue', thickness=3)
            ref_wnd.type_keys('^a' + referanse, with_spaces=True)

            print("OK")                 # Ferdig fylt inn referanse

        if deres_ref:
            print("Fyller inn Deres ref... ", end="")
            your_ref_wnd = (
                self.gui_app.child_window(title_re='(?i)txtYourRef',
                                          control_type='Edit')
                )
            your_ref_wnd.draw_outline(colour='blue', thickness=3)
            your_ref_wnd.type_keys('^a' + deres_ref + '{TAB}',
                                   with_spaces=True)
            print("OK")                 # Ferdig fylt inn Deres ref

        if tab:
            _change_tab(chosen_tab=tab)

        self._save_order()

        # Tilbakefør opprinnelig utklippstavle
        pyperclip.copy(clipboard)
        return order_properties

    def add_orderline(self,
                      prod_num=None,
                      name=None,
                      append_to_name=None,
                      quantity=None,
                      price=None,
                      discount=None,
                      tracking=None
                      ) -> None:
        """Lager én ordrelinje i Mamut.

        Traverserer kolonner og fyller inn verdier som er gitt,
        slutter å traversere dersom det ikke finnes kolonner til høyre
        som skal behandles.
        Forutsetter fokus på 'Produktlinjer'-tabkort i Mamut-ordre

        Parametre:
            prod_num: 'Produktnr.' som settes i Mamut-ordre
            name: 'Beskrivelse' som settes i Mamut-ordre
            append_to_name: tilleggstekst etter 'name' i Mamut-ordre
            quantity: 'Antall' som settes i Mamut-ordre
            price: 'Pris' som settes i Mamut-ordre
            discount: 'Rabatt' som settes i Mamut-ordre
            tracking: 'Sporing' som settes i Mamut-ordre

        Eksempel:
            add_orderline(prod_num='sps330', name='Sensor', discount=24)
        """

        assert prod_num or name, \
            "Minst en av prod_num og name må ha en verdi"

        if prod_num is not None:
            print(f"Fyller inn ordrelinje for {prod_num}... ", end="")
        else:
            print("Fyller inn tekstlinje... ", end="")
        self.gui_app.type_keys('{VK_ADD}')       # '+': ny ordrelinje
        time.sleep(0.35)                    # For sikkerhets skyld

        # NB! Ikke bruk pywinauto for ordrelinjer - bug gjør at Mamut i
        # kombinasjon med pywinauto gjør at radhøyden for ordrelinjen
        # spontant endres og blir enten veldig lav eller veldig høy.

        if prod_num:
            self.wingui.await_text(filltext=prod_num)
            # Gi opp dersom await_text() ikke oppdaterte utklippstavle
            if not pyperclip.paste():
                return None
            # time.sleep(0.2)         # Kanskje nødvendig for å omgå bug
            keyboard.press_and_release('tab')
            self.wingui.await_text()
            # Gi opp dersom await_text() ikke oppdaterte utklippstavle
            if not pyperclip.paste():
                return None


            if append_to_name:
                keyboard.press_and_release('end')
                keyboard.write(append_to_name)
        else:
            keyboard.press_and_release('tab')

        if name is not None:
            self.wingui.await_text(filltext=name)
            # Gi opp dersom await_text() ikke oppdaterte utklippstavle
            if not pyperclip.paste():
                return None

            # time.sleep(0.2)         # Kanskje nødvendig for å omgå bug

        elif name is None and prod_num is None:
            self.wingui.await_text(filltext='.')
            # Gi opp dersom await_text() ikke oppdaterte utklippstavle
            if not pyperclip.paste():
                return None
        if not prod_num:
            # Nødvendig for å omgå bug og sikre oppdatering av felt
            time.sleep(0.2)
            keyboard.press_and_release('tab')
            time.sleep(0.2)

        if tracking:
            keyboard.press_and_release('shift+tab')
            self.wingui.await_text()
            # Gi opp dersom await_text() ikke oppdaterte utklippstavle
            if not pyperclip.paste():
                return None
            keyboard.press_and_release('shift+tab')
            # Gi opp dersom await_text() ikke oppdaterte utklippstavle
            self.wingui.await_text(filltext=tracking)
            if not pyperclip.paste():
                return None
            for __ in range(2):
                keyboard.press_and_release('tab')
                self.wingui.await_text()
                # Gi opp dersom await_text() ikke oppdaterte utklippstavle
                if not pyperclip.paste():
                    return None

        if any([quantity not in [None, 1], price is not None,
                discount is not None]):
            keyboard.press_and_release('tab')
            self.wingui.await_text()
            # Gi opp dersom await_text() ikke oppdaterte utklippstavle
            if not pyperclip.paste():
                return None
            if quantity not in [None, 1]:
                quantity = str(quantity).replace('.', ',')
                self.wingui.await_text(filltext=quantity)
                # Gi opp dersom await_text() ikke oppdaterte utklippstavle
                if not pyperclip.paste():
                    return None
            if any([price is not None, discount is not None]):
                for __ in range(2):
                    keyboard.press_and_release('tab')
                    self.wingui.await_text()
                    # Gi opp dersom await_text() ikke oppdaterte utklippstavle
                    if not pyperclip.paste():
                        return None
                if price is not None:
                    price = str(price).replace('.', ',')
                    self.wingui.await_text(filltext=price)
                    # Gi opp dersom await_text() ikke oppdaterte utklippstavle
                    if not pyperclip.paste():
                        return None
                if discount is not None:
                    discount = str(discount).replace('.', ',')
                    keyboard.press_and_release('tab')
                    self.wingui.await_text(filltext=discount)
                    # Gi opp dersom await_text() ikke oppdaterte utklippstavle
                    if not pyperclip.paste():
                        return None
            # Iblant nødvendig for å sikre oppdatering av linje
            keyboard.press_and_release('tab')
        print("OK")                     # Ferdig utfylt ordrelinje
        return None

    def add_more_orderlines(self, orderlines) -> None:
        """Lager flere ordrelinjer i Mamut.

        Obligatorisk innparameter: dictionary med aktuelle nøkler og
        verdier.
        Eksempel: add_more_orderlines([
            {'prod_num': 'sps340', 'name': 'Sensor', 'price': 5},
            {'prod_num': 'spcc01'}
        ])
        """
        for orderline in orderlines:
            self.add_orderline(**(orderline))
        return None

    def get_serial_number_mask(
        self,
        *, prod_num,
        storage_name='A1: Salgsvarer for salg, internt (Diagnostica)',
    ):
        """
        Lager maske ut fra tilgjengelige serienumre på Mamut-salgslager,
        Obligatorisk innparameter: prod_num (Mamut-produktnummer)
        Eksempel: get_serial_number_mask('SPE200') -> '061xxxxxxx'
        (Alle serienumrene på lager har her sifrene 061 til felles.)
        """

        assert prod_num in SN_PROD_NUMS, \
            f"Ugyldig produktnr: {prod_num}. Tillatte produktnumre: " + \
            ', '.join(SN_PROD_NUMS)

        serial_number_objects = self.lookup_db(f"""
            SELECT g_storeitem.serialnr
            FROM g_storeitem
            JOIN g_store ON g_store.pk_storeid=g_storeitem.fk_store
            JOIN g_prod ON g_prod.pk_prodid = fk_product
            WHERE g_storeitem.outtype IS NULL AND
            g_store.description='{storage_name}'
            AND g_prod.prodid='{prod_num}'
        """)

        if serial_number_objects == [None]:
            serial_number_mask = 'xxx'
        else:
            serial_numbers = [str(e.serialnr) for e in serial_number_objects]
            zipped_serial_numbers = list(zip(*serial_numbers))
            serial_number_mask_list = [set(e).pop() if len(list(set(e))) == 1
                                       else 'x' for e in zipped_serial_numbers]
            serial_number_mask = ''.join(serial_number_mask_list)
        return serial_number_mask
