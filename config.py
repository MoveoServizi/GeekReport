from pathlib import Path
# EMAIL


DESTINATARI = [ #incidenti standard
    "ufficioprogetti@moveoservizi.com",

    "elena.pagani@medicair.it",
    "luca.cozzi@medicair.it",
    "basilio.pau@moveoservizi.com",
]

DESTINATARI_EVENTI_RILIEVO = [ #no incidenti standard
    "michele.cecchi@moveoservizi.com",
    "direzione@moveoservizi.com",
]

# alias opzionale
destinatari = DESTINATARI

#PATH
REPORT_BASE_DIR = Path(r"\\192.168.0.10\Ufficio_Tecnico\PROGETTI\CLIENTI\MEDICAIR\GEEK_VENTILO_1\MANUTENZIONE\REPORT_APP")
REPORT_INFO_IMPIANTO_DIR = Path(r"\\192.168.0.10\Ufficio_Tecnico\PROGETTI\CLIENTI\MEDICAIR\GEEK_VENTILO_1\MANUTENZIONE\REPORT_APP\INFO_IMPIANTO")
LATEX_PATH = Path(r"C:\Users\Administrator\AppData\Local\Programs\MiKTeX\miktex\bin\x64")