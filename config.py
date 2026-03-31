from pathlib import Path
# EMAIL

DESTINATARI = [ #no incidenti standard
   "michele.cecchi@moveoservizi.com",
    "bp.moveo@gmail.com",
    "direzione@moveoservizi.com",
    "elena.pagani@medicair.it",
    "luca.cozzi@medicair.it",
]

DESTINATARI_ALL_REPORT = [ #incidenti standard
    "michele.cecchi@moveoservizi.com",
    "bp.moveo@gmail.com",
]

# alias opzionale
destinatari = DESTINATARI

#PATH
REPORT_BASE_DIR = Path(r"\\192.168.0.10\Ufficio_Tecnico\PROGETTI\CLIENTI\MEDICAIR\GEEK_VENTILO_1\MANUTENZIONE\REPORT_APP")
LATEX_PATH = Path(r"C:\Users\Administrator\AppData\Local\Programs\MiKTeX\miktex\bin\x64")