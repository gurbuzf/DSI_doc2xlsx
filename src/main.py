############################
# Written by: Faruk Gurbuz #
# 26.04.2022 MIT License   #
############################

import os, sys
from datetime import datetime
from dsi_doc2xlsx import read_DSI_doc_file
from rich import print
from rich.padding import Padding
from rich.console import Console
from rich.text import Text
from rich.style import Style
from rich.markdown import Markdown
from rich.prompt import Prompt

from rich.progress import (
    BarColumn,
    DownloadColumn,
    Progress,
    TaskID,
    TextColumn,
    TimeRemainingColumn,
    TransferSpeedColumn,
)

progress = Progress(
    TextColumn("[bold blue]{task.fields[filename]}", justify="right"),
    BarColumn(bar_width=None),
    "[progress.percentage]{task.percentage:>3.1f}%",
    "•",
    DownloadColumn(),
    "•",
    TransferSpeedColumn(),
    "•",
    TimeRemainingColumn()
)

def read_write_doc(file, path2doc, path2save, t_id:TaskID):
    name = file.split('.')[0]
    _path = os.path.join(path2doc, file)
    temp = read_DSI_doc_file(_path)
    path_ = os.path.join(path2save, f'{name}.xlsx')
    temp.write_xlsx(path_)
    progress.update(t_id, advance=1)
    progress.console.log(f"Kaydedildi| {path_}")

def progress_writing(files, path2doc, path2save):
    with progress:
        for file in files:
            t_id= progress.add_task('[cyan]Çalışıyor...', filename=file, total=len(files))
            read_write_doc(file, path2doc, path2save, t_id)

def prompt_create_directory(path2create):  
    """ Yeni Klasör oluşturmak için kullanıcı girdisi alır.    
    """
      
    name = Prompt.ask(f"\n[bold white]{path2create}[/bold white] [bold red]dizinini oluştur[evet-hayır].")
    approve = ['evet', 'e', 'yes', 'y']
    approve = approve + [i.upper() for i in approve]
    disapprove = ['hayır', 'h', 'no', 'n']
    disapprove = disapprove + [i.upper() for i in disapprove]
    if name in approve:
        try:
            os.mkdir(path2create)
            print(f"[+] {path2create} [bold green]dizini oluşturuldu!")
        except OSError as error:
            print("HATA:" + error) 
            sys.exit(f"[+] {path2create} [bold red]dizini oluşturulamadı! Program sonlandırılıyor!")           
    elif name in disapprove:
        sys.exit("[bold red]Program sonlandırılıyor![/bold red]")
    else:
        print("Lütfen geçerli bir seçim yapınız! ['evet' 'e' 'yes' 'y' 'hayır' 'h' 'no' 'n']")
        prompt_create_directory(path2create) 

def main():
    console = Console()
    ######Title
    style = Style(color="purple4", bgcolor="black", bold=True)
    text = Text("==============AÇIKLAMA==============\n \
                Bu program .doc formatındaki akım yıllıklarını .xlsx formatına dönüştürmek için hazırlanmıştır.", justify="center")
    txt1 = Padding(text, (1, 1), style=style, expand=True)
    print(txt1)
    
    ######Readme
    MARKDOWN = """
    1. Dönüştürülecek .doc dosyalarını bir klasör içine atınız. [.doc] uzantılı dosyaların bulunduğu klasör dizinini (dosya yolu) DİZİN-oku seçeneğine gir. 
    Örnek:[DİZİN-oku: C:\\Users\\farukgurbuz\\E13A059]

    2. [.xlsx] dosyalarının kaydedileceği klasör dizinini [DİZİN-yaz] seçeneğine gir. Eğer belirtilen isimde bir klasör yoksa klasör oluşturulması için evet/hayır seçeneğine evet yazılmalıdır. 
    Diğer geçerli secenekler > ['evet' 'e' 'yes' 'y' 'hayır' 'h' 'no' 'n']

    3. Dönüştürme işlemi tamamlandığında [TAMAMLANDI] bilgisi ekrana gelir ve pencere kapanır.
    """
    md = Markdown(MARKDOWN)
    console.print(md)

    #####Footnote
    text = Text("Devlet Su İşleri Genel Müdürlüğü 2022\u00a9", justify="center")
    txt2 = Padding(text, (0, 0), 
                 style=style, expand=True)
    print(txt2)
    ##########################################

    #####File/Folder management
    path2doc = console.input("[bold green]1- Lütfen .doc uzantılı dosyaların bulunduğu klasörün dosya yolunu giriniz! \n\n • DİZİN-oku:")
    if os.path.isdir(path2doc):       
        pass
    else:
        print("[bold red]UYARI![/bold red] Dosya yolu bulunamadı. Program sonlandırdı!")
        sys.exit(f"{path2doc}")
    
    path2save = console.input("\n[bold magenta]2-Lütfen .xlsx dosyalarını kaydetmek istediğiniz klasörün dosya yolunu giriniz! \n\n • DİZİN-yaz:[/bold magenta]")
    
    if os.path.isdir(path2save):       
        pass
    else:
        print("[bold red]UYARI![/bold red] Dosya yolu bulunamadı!")
        prompt_create_directory(path2save)  
    ##########################################

    ######Transformation        
    files = [d for d in os.listdir(path2doc) if os.path.isfile(os.path.join(path2doc, d)) and d.endswith('doc')]
    print("\n[yellow] BAŞLADI  Saat:" + datetime.now().strftime('%d/%m/%Y %H:%M:%S'))

    print(f'[bold yellow][+][/bold yellow] {path2doc} dizinindeki dosyalar okunuyor...')
    progress_writing(files, path2doc, path2save)

    print("\n[cyan] TAMAMLANDI  Saat:" + datetime.now().strftime('%d/%m/%Y %H:%M:%S'))


if __name__ == "__main__":
    main()