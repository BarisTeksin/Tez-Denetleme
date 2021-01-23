from docx import *
import docx.package
import docx.parts.document
import docx.parts.numbering
import logging

logging.basicConfig(filename="tez.log", format='%(asctime)s %(message)s', filemode='w') 
logger=logging.getLogger()
logger.setLevel(logging.DEBUG)


def Kaynakca_Kontrol(kaynakca,paragraphs):
    kaynakca_pass = ['ve','pp.','ss.','of','by','and','in','the','cilt']
    for satir,kaynakca_satir in enumerate(kaynakca,1):
        kaynakca_no = int(kaynakca_satir.split('[')[1].split(']')[0])
        if satir != kaynakca_no:
            logging.info('{} Numaralı kaynakça bulunamadı.'.format(satir))
        for kelime in kaynakca_satir.split(' '):
            if kelime[0] != '(' and not (kelime.replace('[{}]'.format(satir),'').strip()[0].isupper() or kelime.replace('[{}]'.format(satir),'').strip()[0].isnumeric() or kelime in kaynakca_pass):
                logging.info('{} Numaralı kaynakçada format hatası. {} kelimesi küçük harf'.format(kaynakca_no,kelime))
        kontrol = True
        for paragraph in paragraphs:
            if 'Kaynaklar' in paragraph.text:
                break
            if '[{}]'.format(kaynakca_no) in paragraph.text:
                kontrol = False
                break
        if kontrol:
            logging.info('{} Numaralı kaynakça tezde bulunamadı.'.format(kaynakca_no))


def main(filename):
    document = Document(filename+'.docx')
    paragraphs = document.paragraphs
    kontrol = False
    kaynakca_sayisi_toplam = 0
    kaynakca_listesi = []
    onsoz_sayi = 0
    for sayi,paragraph in enumerate(paragraphs):
        if 'Kaynaklar' in paragraph.text:
            kontrol = True
        elif 'Özgeçmiş' == paragraph.text:
            kontrol = False
        elif 'ÖN SÖZ' == paragraph.text:
            onsoz_sayi = sayi
        if kontrol and paragraph.text != '' and '[' in paragraph.text:
            kaynakca_listesi.append(paragraph.text.strip())
        elif sayi == onsoz_sayi + 1:
            if 'teşekkür' in paragraph.text.lower():
                logging.info('Önsözün ilk paragrafında "teşekkür" ifadesi mevcut.')
        if '“' in paragraph.text:
            if len(paragraph.text.split('“')[1].split('”')[0].split(' ')) > 50:
                logging.info('{} ile başlayan paragraf {} kelime.'.format(paragraph.text.split(' ')[0] + ' ' + paragraph.text.split(' ')[1] + ' ' +paragraph.text.split(' ')[2],len(paragraph.text.split('“')[1].split('”')[0].split(' '))))

    Kaynakca_Kontrol(kaynakca_listesi,paragraphs)

if __name__ == "__main__":
    filename = input('Tez dosyasının adı : ')
    main(filename)