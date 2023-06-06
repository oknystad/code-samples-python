# -*- encoding: utf-8 -*-#
# !/usr/bin/python

# Standardbibliotek import
import lxml.builder
import lxml.etree
import os
import subprocess
import time
import traceback
from typing import Optional
import win32print

# Tredjeparts bibliotek import
import pendulum
import pyperclip
from pywinauto.application import Application as PWA_App

# Lokal applikasjon import
import okn_functions as okn
from okn_console_function import console_setting
from okn_constants import MAMULARE_RE, PROGRAM_PATH
from okn_ext_classes import MenuMaker, WinGUIManager

__author__ = 'Øyvind Nystad'


wingui = WinGUIManager()


def generate_xml_content(*, text_lines = ' ',
                         tape_width: int = 12,
                         is_border: bool = True) -> None:

    def _normalize_text_lengths(text_lines: list) -> list:
        # Forleng alle tekstlinjer til lengste tekstrengs lengde
        cell_len = len(max(text_lines, key=len))

        # Tegn som krever to plasser på RHINO etikett, men som
        # gir len() lik 1 - må justeres for.
        chubby_chars = '⚠→'

        normalized_text_lines = []
        for text_line in text_lines:
            count_chubby_chars = 0
            for s in text_line:
                if s in chubby_chars:
                    count_chubby_chars += 1
            normalized_text_lines.append(
                text_line.ljust(cell_len - count_chubby_chars))
        return normalized_text_lines

    text_lines = _normalize_text_lengths(text_lines)

    E = lxml.builder.ElementMaker()
    PLIST = E.plist
    DICT = E.dict
    KEY = E.key
    STRING = E.string
    INTEGER = E.integer
    FALSE = E.false
    REAL = E.real
    TRUE = E.true
    ARRAY = E.array

    def _dymo_celldata(text_line: str) -> DICT:
        return DICT(
                   KEY('Text'),
                   STRING(text_line),
                   KEY('ClassName'),
                   STRING('DYMO.RhinoPro.SDK.CellDataText'),
                )
    args = []
    for text_line in text_lines[:-1]:
        args.append(_dymo_celldata(text_line))
        # Legg til linjeseparator
        args.append(DICT(
            KEY('ClassName'),
            STRING('DYMO.RhinoPro.SDK.CellDataLineSeparator'),
            )
        )
    # Legg til siste tekstlinje
    args.append(_dymo_celldata(text_lines[-1]))

    xml_structure = PLIST(
        DICT(
            KEY('LabelTemplate'),
            DICT(
                KEY('TapeWidth'), STRING(f'tw{tape_width}'),
                KEY('Font'),
                DICT(
                    # Inconsolata er brukt bl.a. fordi den er
                    # en Monospace-font
                    KEY('FontFamily'), STRING('Inconsolata'),
                    KEY('FontSize'), REAL('48'),
                    KEY('Charset'), INTEGER('1'),
                    KEY('Bold'), TRUE(),
                ),
                KEY('ShrinkToFit'), TRUE(),
                KEY('CutMode'), STRING('AutoCut'),
                KEY('Border'), TRUE() if is_border else FALSE(),
                KEY('DeviceFont'),
                DICT(
                    KEY('FontType'), STRING('Arial'),
                    KEY('FontSize'), STRING('XXL'),
                    KEY('FontStyle'), STRING('Regular'),
                    ),
                KEY('FixedLabelLength'), REAL('5040'),
                KEY('CellsData'), ARRAY(DICT()),
                KEY('TemplateClass'), STRING('DYMO.RhinoPro.SDK.GeneralLabel'),
                ),
            KEY('Labeldata'),
            DICT(
                KEY('Rows'),
                ARRAY(ARRAY(DICT(
                    KEY('Items'), ARRAY(*args)
                      )
                    )
                  )
                )
              )
            )
    xml_code = (lxml.etree
                    .tostring(xml_structure, pretty_print=True)
                    .decode())
    return xml_code


def get_mamulare_cust_info() -> tuple[int, str, str]:
    """Hent kundeinformasjon fra Mamulare."""
    wingui.wnd_focus(title_re=MAMULARE_RE, is_maximized=True)
    try:
        mamulare_app = (PWA_App(backend="uia")
                        .connect(title_re=MAMULARE_RE, found_index=0)
                        .window(title_re=MAMULARE_RE))
    except Exception:
        print("Mamulare må være åpen. Avslutter...")
        time.sleep(3.)

    # Åpne Lisensinformasjon-vindu
    (mamulare_app.child_window(title='Rediger lisensinformasjon')
                 .type_keys('{ENTER}'))

    cust_num = (mamulare_app
                .child_window(found_index=5,
                              class_name='TcxCustomInnerTextEdit')
                .iface_value.CurrentValue)

    version = (mamulare_app
               .child_window(found_index=0,
                             class_name='TcxCustomInnerTextEdit')
               .iface_value.CurrentValue)

    expiry_nor = (mamulare_app
                  .child_window(found_index=0,
                                class_name='TcxCustomDropDownInnerEdit')
                  .iface_value.CurrentValue)

    # Lukk Lisensinformasjon-vindu
    (mamulare_app.child_window(found_index=0, title='Avbryt')
                 .type_keys('{ENTER}'))
    return cust_num, version, expiry_nor


def _create_rhino_file(xml_code: str) -> None:
    with open(fr'{PROGRAM_PATH}\Temp\label_file.rhino', 'w') as text_file:
        text_file.write(xml_code)
    return None


def _open_rhino_file() -> None:
    """Åpne etikettfil i RHINO Connect."""

    # Lukk RHINO Connect "forcefully" dersom åpent
    print("Åpner RHINO Connect med valgt etikett... ", end="")
    subprocess.call(['taskkill', '/F', '/IM', 'RhinoPro.exe'],
                    stdout=open(os.devnull, 'w'),
                    stderr=subprocess.STDOUT,
                    )
    okn.start_winprog(name='rhino_connect',
                      target=fr'{PROGRAM_PATH}\Temp\label_file.rhino')
    wingui.wnd_focus(title_re='.*RHINO.*', timeout=10.)
    print("OK")         # Åpnet Rhino Connect med etikett
    return None


def get_chosen_label_vals()-> Optional[tuple[list, int, bool, int]]:
    """Hent verdier for etikett."""
    print()
    label_menu = MenuMaker(
        'TYPE LABEL',
        ('1', 'Kalibrering med dato, engelsk, 19 mm'),
        ('2', 'Batteriinfo med dato, norsk, 12 mm'),
    )
    label_opt = label_menu().key

    cust_num = ''
    version = ''
    expiry_nor = ''
    expiry_eng_swe = ''
    if label_opt in ['6', '7', '8']:
        cust_num, version, expiry_nor = get_mamulare_cust_info()
        # Fra format DD.MM.YYYY til YYYY-MM-DD
        expiry_eng_swe = (pendulum.from_format(expiry_nor, 'DD.MM.YYYY')
                                  .strftime("%Y-%m-%d"))

    label_values = {    # text_lines, tape_width, is_border
        '1': [['   CALIBRATION    │     ⚠',
               'By:   [COMPANY]  '
               " │ Don't tug/",
               f'Date: {pendulum.now().strftime("%Y-%m-%d")}'
               '  │ twist when',
               f'Due:  {pendulum.now().add(years=2).strftime("%Y-%m-%d")}'
               '  │ unplugging',
               ' www.[COMPANY].com  │ cable   →'], 19, True],
        '2': [['DIAGNOSTICA AS    Varighet:',
               f'Dato: {pendulum.now().strftime("%d.%m.%Y")}  ca. 1 år /',
               'www.[COMPANY].com   250 ladinger'], 12, False],
        }

    # 4 klistremerker hvis batteri, ellers 1
    print_count = 4 if label_opt in ['2',] else 1

    if label_opt == 'X':
        okn.mention_return_to_main_menu(message="Ingen handling",
                                        duration=1.0)
        return None
    else:
        text_lines, tape_width, is_border = label_values[label_opt]
        return text_lines, tape_width, is_border, print_count


def _get_rhino_printer_awake_status() -> bool:
    """Se til at RHINO 6000-printer er tilgjengelig på PC."""
    is_awake = False
    CODE_ACTIVE_PRINTER = 2624
    CODE_INACTIVE_PRINTER = 3648

    print("Ser etter RHINO 6000-printer på PC-en... ", end="")
    try:
        printer_handle = win32print.OpenPrinter('RHINO 6000')
    except Exception:
        print("feilet")
        okn.mention_return_to_main_menu(
            message="Printeren er ikke installert",
            duration=10.)
        return is_awake
    print("OK")

    # Test om RHINO 6000 er påslått
    status_code = win32print.GetPrinter(printer_handle, 5)['Attributes']

    is_awake = {CODE_ACTIVE_PRINTER: True,
                CODE_INACTIVE_PRINTER: False}[status_code]

    print("Finner ut om printer er våken... ", end="")
    print("Ja") if is_awake else print("Nei")
    if not is_awake:
        okn.mention_return_to_main_menu(
            message="RHINO 6000 er avslått, slå denne på før utskrift\n",
            duration=6.)
    return is_awake


def _open_print_menu(print_count: int = 0) -> None:
    (PWA_App(backend="uia")
     .connect(title_re='.*RHINO Connect Software')
     .window(title_re='.*RHINO Connect Software')
     .type_keys('^p'                    # Åpne print-meny
                '{TAB}'                 # Fokus på 'Number of copies'felt'
                '^a'                    # Marker alt i felt
                f'{print_count}'        # Antall kopier
                '^a'                    # Marker alt i felt på nytt
                )
     )

    okn.mention_return_to_main_menu(message="Trykk ENTER for å skrive ut\n",
                                    duration=6.)
    return None


def main() -> None:
    """Hovedfunksjon."""
    label_vals = get_chosen_label_vals()
    if label_vals is not None:
        text_lines, tape_width, is_border, print_count = label_vals
        console_setting(state='busy')
        xml_code = generate_xml_content(text_lines=text_lines,
                                        tape_width=tape_width,
                                        is_border=is_border)
        _create_rhino_file(xml_code)
        _open_rhino_file()
        is_rhino_printer_awake = _get_rhino_printer_awake_status()
        if is_rhino_printer_awake:
            _open_print_menu(print_count)
    return None


if __name__ == '__main__':
    try:
        main()
    except Exception as ex:
        tb_lines = traceback.format_exception(ex.__class__, ex,
                                              ex.__traceback__)
        for line in tb_lines:
            print(line)
        pyperclip.copy('\r\n'.join(tb_lines))
        console_setting(state='failure')
        os.system('pause')
