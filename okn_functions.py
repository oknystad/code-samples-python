# -*- encoding: utf-8 -*-#
# !/usr/bin/python

# Standardbibliotek import
import bisect
import io
import itertools
import os
import subprocess
import time

# Tredjeparts bibliotek import
import cv2
import selenium.webdriver
import selenium.common
# import win32gui
import win32console

# Lokal applikasjon import
from okn_console_function import console_setting
from okn_constants import PROGRAM_PATH


__author__ = 'Øyvind Nystad'

"""
Egenlagde funksjoner for ofte brukt/nyttig funksjonalitet i Windows:
- default_input         Som input(), med mulighet for standardverdi
- draw_red_dot          Tegn rød prikk, f.eks. ved museklikk
- draw_red_frame        Tegn rød ramme, f.eks. ved bildesøk på skjerm
- find_image            Let etter angitt .png-bilde på skjermen
- start_winprog         Start Windows-program med evt. param. og målfil
- verbose_print         Fyll inn farget tekst uten påfølgende linjeskift
"""



def default_input(prompt: str = '', default_text: str = '') -> str:
    """
    Venter på verdi fra bruker på samme måte som input(),
    men i tillegg kan egendefinert standardverdi settes.

    Args:
        prompt (str): Beskjed foran skrivefelt
        default_text (str): Forhåndsutfylt tekst

    Returns:
        str: Tekst utfylt av bruker
    """

    console_setting(state='ready')
    keys = []
    for c in str(default_text):
        evt = win32console.PyINPUT_RECORDType(win32console.KEY_EVENT)
        evt.Char = c
        evt.RepeatCount = 1
        evt.KeyDown = True
        keys.append(evt)
    stdin = win32console.GetStdHandle(win32console.STD_INPUT_HANDLE)
    stdin.WriteConsoleInput(keys)
    return input(prompt)


def mention_return_to_main_menu(message: str = '',
                                duration = 3.0) -> None:
    """
    Skriv beskjed om at man returnerer til hovedmeny.
    Sett

    Args:
        message (str): Tekstbeskjed før retur til hovedmeny begynner
        duration (float): Varighet på beskjeden
    Returns:
        None
    """

    console_setting(state='busy')
    if message:
        print(message)
    print("Returnerer til hovedmeny", end="")
    # Beskjed vises i 3 sekunder
    DOT_COUNT = 10
    for __ in range(DOT_COUNT):
        print(".", end="")
        time.sleep(duration / DOT_COUNT)
    return None


def is_image_inside_image(small_image_path: str,
                          large_image_path: str) -> bool:
    """
    Fastslå om et lite .png-bilde er del av et større bilde.

    Args:
        small_image_path (str): Sti til lite bilde
        large_image (str): sti til stort bilde

    Returns:
        bool: True hvis stort bilde inneholder lite bilde, ellers False
    """

    ERROR_MARGIN = 0.05
    small_image_obj = cv2.imread(small_image_path)
    large_image_obj = cv2.imread(large_image_path)

    result = cv2.matchTemplate(
        small_image_obj,
        large_image_obj,
        method=cv2.TM_SQDIFF_NORMED)

    min_squared_diff = cv2.minMaxLoc(result)[0]

    return min_squared_diff < ERROR_MARGIN


def start_winprog(*, name = None, param = '',
                  target = '') -> None:

    """
    Starter opp Windows-program. Mulige innparametere:
    - param: Overstyrer def_param (standard-parameter)
    - target: mål, f.eks fil for å åpnes eller nettsted
    """

    def find_valid_prog_file_path(possible_prog_file_paths):
        """
        Finn sti for program på brukerens PC -
        Windows-sti kan variere fra PC til PC
        """
        valid_prog_file_paths = [
            f for f in possible_prog_file_paths if os.path.isfile(f)]
        assert valid_prog_file_paths, (
            "Klarte ikke å starte program - ingen av følgende stier er "
            "gyldige:\n"
            fr"{str(possible_prog_file_paths)}"
        )

        valid_prog_file_path = valid_prog_file_paths[0]
        return valid_prog_file_path

    acrobat_prog_path = find_valid_prog_file_path(
        possible_prog_file_paths=(
        r'C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe',
        r'C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe',
        )
    )

    outlook_prog_path = find_valid_prog_file_path(
        possible_prog_file_paths=(
        r'C:\Program Files\Microsoft Office\root\Office16\outlook.exe',
        r'C:\Program Files (x86)\Microsoft Office\root\Office16\outlook.exe',
        )
    )

    win_progs = {   # [navn]: ([parameter], [sti]),
        'acrobat': ('', acrobat_prog_path),
        'firefox': ('', r'C:\Program Files\Mozilla Firefox\firefox.exe'),
        'mamut': ('', r'C:\Program Files (x86)\Mamut\Mamut.exe'),
        'outlook': ('', outlook_prog_path),
        'phonero': ('', r'Z:\CDB\Phonero\Phonero.exe'),
        'rhino_connect': ('', r'C:\Program Files (x86)\RHINO Connect '
                          r'Software\RhinoPro.exe'),
        'word': ('', r'C:\Program Files (x86)\Microsoft Office'
                 r'\root\Office16\winword.exe'),
    }
    path = os.path.dirname(win_progs[name][1])
    executable = os.path.basename(win_progs[name][1])
    os.chdir(path)

    run_command = fr'"{executable}" {param}'
    run_command += fr' "{target}"' if target else ''

    subprocess.Popen(run_command)
    return None
