# -*- encoding: utf-8 -*-#
# !/usr/bin/python

# Standardbibliotek import
import copy
import ctypes
import msvcrt
import pyperclip
import re
import time

# Tredjeparts bibliotek import
import colorama
import win32api
import win32con
import win32gui
import win32com
import win32com.client

# Lokal applikasjon import
from okn_console_function import console_setting
from okn_constants import PROGRAM_RE

__author__ = 'Øyvind Nystad'

"""
Egenlagde klasser for grunnleggende Adviuvare-funksjoner
- NamedList
- BasicWndHandler
- MenuMaker
"""


class NamedList():
    """Liste over nøkler med verdier, overskrivbart innhold.

    Virkemåte lik types.SimpleNamespace, men med flere metoder.

    Metoder:
        __len__: Returnerer antall listeelementer
        update: Slår NamedList sammen med NamedList eller dict
            og/eller kwargs.

    Eksempel:
        foo = NamedList(x=1, y=2)
        >>> print(foo.x)
        >>> 1
    """

    def __init__(self, **kwargs) -> None:
        """Initialiserer NamedList"""
        self.__dict__.update(kwargs)

    def __repr__(self) -> str:
        keys = sorted(self.__dict__)
        items = (f'{k}={self.__dict__[k]}' for k in keys)
        return f'{type(self).__name__}({", ".join(items)})'

    def __eq__(self, other) -> bool:
        if hasattr(other, '__dict__'):
            return self.__dict__ == other.__dict__
        else:
            return False

    def __len__(self) -> int:
        return len(self.__dict__)

    def update(self, other=None, **kwargs):
        """Oppdater en NamedList

        NameList slås sammen med verdier fra annen NamedList eller
        dict, og/eller kwargs. Ved lik nøkkel som eksisterende,
        overskrives verdi.

        Eksempel:
            >>> foo = NamedList(a=1, b=2)
            >>> bar = NamedList(b=3, c=4)
            >>> baz = foo.update(bar)
            >>> print(baz)
            NamedList(a=1, b=3, c=4)
            >>> baz = foo.update({'d': 5, 'e': 6})
            NamedList(a=1, b=2, d=5, e=6)
            >>> baz = foo.update(f=7)
            NamedList(a=1, b=2, f=7)
        """
        merged = copy.copy(self)
        if other and type(self) == type(other):
            for key, value in other.__dict__.items():
                merged.__dict__[key] = value
        elif isinstance(other, dict):
            for key, value in other.items():
                merged.__dict__[key] = value
        merged.__dict__.update(kwargs)
        return merged


class BasicWndHandler:
    """Klasse for enkel manipulering av vinduer i Windows.

    :param x_pixel_res: Skjermoppløsning i X-retning
    :param y_pixel_res: Skjermoppløsning i Y-retning

    Metoder:
        get_active_wnds: Finn alle synlige og aktive vinduer.
        wnd_focus: Sett fokus på ønsket vindu.
    """

    def __init__(self) -> None:
        """Finn skjermens oppløsning."""
        self.x_pixel_res: int = win32api.GetSystemMetrics(0)
        self.y_pixel_res: int = win32api.GetSystemMetrics(1)

    def _get_all_wnd_handles(self) -> list[int]:
        """Returner handle-id for alle vinduer.

        Opplistingen er i z-rekkefølge, dermed også i kronologisk
        rekkefølge etter når vinduene sist var i fokus.
        """
        wnd_handles: list[int] = []
        top_wnd_handle: int = ctypes.windll.user32.GetTopWindow(None)
        if top_wnd_handle:
            wnd_handles.append(top_wnd_handle)
            while True:
                next_wnd_handle: int = ctypes.windll.user32.GetWindow(
                    wnd_handles[-1], win32con.GW_HWNDNEXT)
                if not next_wnd_handle:
                    break
                wnd_handles.append(next_wnd_handle)
        return wnd_handles

    def get_focused_wnd_title(self):
        """Finn tittel for fokusert vindu."""

        wnd_handle = ctypes.windll.user32.GetForegroundWindow()
        return win32gui.GetWindowText(wnd_handle)

    def get_active_wnds(self) -> list[tuple[int, int, str]]:
        """Finn alle synlige vinduer.

        Returner z-nummer (nummer i rekkefølge, fra toppvindu og
        nedover, indikerer kronologi), handle-id og tittel for
        hvert vindu.
        """
        all_wnd_handles: list[int] = self._get_all_wnd_handles()
        active_wnd_handles: list[int] = [
            wnd_handle for wnd_handle in all_wnd_handles if
            win32gui.IsWindowVisible(wnd_handle) and
            win32gui.GetWindowText(wnd_handle)]
        active_wnds: List[Tuple[int, int, str]] = [
            (idx, wnd_handle, win32gui.GetWindowText(wnd_handle)) for
            idx, wnd_handle in enumerate(active_wnd_handles)]
        return(active_wnds)

    def _get_wnd_match(self,
                       is_case_sensitive,
                       timeout: float,
                       title_re: str,
                       is_verbose) -> int | None:
        """Finn ønsket vindu."""
        init_time = time.time()
        wnd_matches: list = []

        while not wnd_matches and time.time() - init_time < timeout:
            active_wnds = self.get_active_wnds()

            for wnd in active_wnds:
                if is_case_sensitive:
                    if re.match(title_re, wnd[2]):
                        wnd_matches.append(wnd)
                else:
                    if re.match(title_re, wnd[2], re.IGNORECASE):
                        wnd_matches.append(wnd)
        if wnd_matches:

            wnd_handle = wnd_matches[0][1]
            wnd_title = wnd_matches[0][2]
            if is_verbose:
                print(f"Fant vindu {wnd_handle} - '{wnd_title}'")
        else:
            if is_verbose:
                print("Fant ingen vinduer med tittel som matchet\n"
                      f"mønsteret {title_re} etter {timeout} sekunder")
            wnd_handle = None
            wnd_title = None
        return wnd_handle, wnd_title

    def _get_updated_wnd_vals(self, wnd_handle, x, y, w, h):
        x_left, y_upper, x_right, y_lower = win32gui.GetWindowRect(wnd_handle)

        # Beregning av ny w
        if isinstance(w, int):
            w_new = w
        elif isinstance(w, float):
            w_new = int(self.x_pixel_res * w / 100.)
        elif w is None:
            w_new = x_right - x_left

        # Beregning av ny h
        if isinstance(h, int):
            h_new = h
        elif isinstance(h, float):
            h_new = int(self.y_pixel_res * h / 100.)
        elif h is None:
            h_new = y_lower - y_upper

        # Beregning av ny x
        if isinstance(x, int):
            if not -self.x_pixel_res <= x <= self.x_pixel_res:
                raise ValueError(f"x={x} utenfor gyldig område "
                                 f"[{-self.x_pixel_res}, {self.x_pixel_res}]")
            if x < 0:
                x_new = self.x_pixel_res - w_new + x
            else:
                x_new = x
        elif isinstance(x, float):
            if not 0. <= x <= 100.:
                raise ValueError(f"x={x} utenfor gyldig område "
                                 f"[0., 100.]")
            x_new = int(self.x_pixel_res * x / 100.)
        elif x is None:
            x_new = x_left

        # Beregning av ny y
        if isinstance(y, int):
            assert abs(y) < self.y_pixel_res, \
                (f"y={y} utenfor gyldig område "
                 fr"[-{self.y_pixel_res}, {self.y_pixel_res}]")
            if y < 0:
                y_new = self.y_pixel_res - h_new + y
            else:
                y_new = y
        elif isinstance(y, float):
            assert abs(y) < 100., \
                f"y={y} utenfor gyldig område [-100, 100]"
            y_new = int(self.y_pixel_res * y / 100.)
        elif y is None:
            y_new = y_upper

        return x_new, y_new, w_new, h_new

    def wnd_focus(self, *,
                  is_case_sensitive: bool = True,
                  is_maximized: bool = False,
                  timeout: float = 2.0,
                  title_re: str = '.*',
                  is_topmost: bool = False,
                  is_verbose: bool = False,
                  x: int | float | None = None,
                  y: int | float | None = None,
                  w: int | float | None = None,
                  h: int | float | None = None,
                  ) -> dict | None:

        """
        Sett fokus på ønsket vindu.
        Dersom flere treff, velges vindu med lavest z-verdi,
        dvs. vinduet som er øverst og dermed var sist i fokus.

        :param is_case_sensitive: Angi om store og små bokstaver skal
            tas hensyn til ved regex-søk
        :param is_maximized: Angi om vindu skal fylle hele skjermen.
            Hvis både is_maximized=True, og en av w, y, w og h har
            verdi, vil vinduet vises som maksimert, og definert
            størrelse vil ikke synes før vinduet minimeres igjen.
        :param timeout: Tidsperiode før forsøk på å sette fokus gis opp
        :param title_re: Regex som identifiserer vinduets tittel
        :param is_topmost: Angi om vindu alltid skal være synlig
        :param is_verbose: Angi om det skal gis beskjed i terminalvindu
            om resultatet av vindussøk
        :param x: X-koordinatverdi som settes for vinduets venstre kant
        :param y: Y-koordinatverdi som settes for vinduets øvre kant
        :param w: Ønsket vindusbredde
        :param h: Ønsket vindushøyde

        :return: dict med elementer wnd_handle og wnd_title

        Eksempel:
            # Sett fokus på Notisblokk, og sett bredde og høyde til
            # 50 % av skjermens oppløsning, med plassering av venstre
            # øvre hjørne i koordinater (100, 100).
            foo = BasicWndHandler()
            foo.wnd_focus(title_re=r'.*Notisblokk.*', is_maximized=True,
                          x=100, y=100, h=50., w=50.)
        """

        # Parenteser inni regex-uttrykk må escapes med \-tegn
        title_re = title_re.replace('(', r'\(')
        title_re = title_re.replace(')', r'\)')

        wnd_handle, wnd_title = self._get_wnd_match(is_case_sensitive, timeout,
                                              title_re, is_verbose)

        if wnd_handle:
            # Gjør vindu synlig
            win32gui.ShowWindow(
                wnd_handle,
                win32con.SW_MAXIMIZE if is_maximized else   # 3
                win32con.SW_NORMAL)                         # 1

            if is_topmost:                 # Vindu alltid øverst
                win32gui.SetWindowPos(
                    wnd_handle, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                    win32con.SWP_NOSIZE | win32con.SWP_NOMOVE)

            if any([x, y, w, h]):
                x_new, y_new, w_new, h_new = self._get_updated_wnd_vals(
                    wnd_handle, x, y, w, h)
                win32gui.MoveWindow(wnd_handle, x_new, y_new, w_new, h_new, 0)

            # Gir vindu fokus
            try:
                win32gui.SetForegroundWindow(wnd_handle)
            except Exception:           # pywintypes.error
                # Iblant feiler win32gui.SetForegroundWindow(). Fiks fra
                # https://stackoverflow.com/a/15503675 - send først ett
                # Alt-knappetrykk til vindu.
                win32com.client.Dispatch('WScript.Shell').SendKeys('%')
                win32gui.SetForegroundWindow(wnd_handle)
            # Sannsynligvis nødvendig pause, evt. juster opp dersom
            # pywinauto.findwindows.ElementNotFoundError oppstår
            time.sleep(0.3)

        return dict(wnd_handle=wnd_handle,
                    wnd_title=wnd_title)


class MenuMaker:
    def __init__(self,
                 title=None,
                 *args,
                 ):
        """Initier valgmeny.

        Argumenter:
            title: Menytittel, vil stå øverst uthevet
            args: tuple med menyelementer, som hver kan være
                en tekststreng for informasjon, eller tuple med
                lengde 2, som inneholder hurtigtast, og
                funksjonsbeskrivelse.
        """
        self.title = title
        self.args = args + (('X', 'Avbryt'),)
        self.wndhandler = BasicWndHandler()

    def _shout(self, s):
        """Returner tekststreng i farge og caps lock.

        colorama.init må være aktivert for at farger skal vises.
        """
        return (colorama.Back.RED + colorama.Style.DIM +
                colorama.Fore.WHITE + s.upper() +
                colorama.Style.RESET_ALL)

    def __call__(self):
        colorama.init(autoreset=True)

        if self.title:
            print('   ', self._shout(self.title))
        option_by_key = {}
        for elem in self.args:
            if isinstance(elem, str):
                print('    ' + elem.replace('\n', '\n    '))
            elif isinstance(elem, tuple):
                key, descr = str(elem[0]), str(elem[1])
                option_by_key[key.upper()] = descr
                pattern = re.compile(key, re.IGNORECASE)
                option_line = pattern.sub(self._shout(key), descr, 1)
                print(self._shout(f'[{key}]') + ' ' + option_line)

        console_setting(state='ready')
        self.wndhandler.wnd_focus(title_re=PROGRAM_RE)
        pressed_key = None

        while True:
            pressed_key = str(msvcrt.getch())[2].upper()
            if pressed_key in option_by_key.keys():
                break
            else:
                print(self._shout("Ugyldig valg!"), end="")
                time.sleep(0.3)
                print("\r" + " " * 30, end="\r")

        self.wndhandler.wnd_focus(title_re=PROGRAM_RE,
                                  x=-25,
                                  y=-90,
                                  w=410,
                                  h=440,
                                  is_maximized=False,
                                  is_topmost=True)

        console_setting(state='busy')

        chosen_option = option_by_key[pressed_key]
        print()
        print(self._shout(f'=> {chosen_option}'))
        return NamedList(
            key=pressed_key,
            option_desc=chosen_option,
        )

    def __repr__(self):
        return f'{__class__.__name__}{(self.title,) + self.args[:-1]}'
