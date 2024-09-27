import pyodbc
import os
import sys
from docxtpl import DocxTemplate
from datetime import datetime as dt
import win32com.client as win32
from PyPDF2 import PdfMerger
import PyPDF2
import docx2pdf
import shutil

pagina_goala = "G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/DOCUMENTE/pagina_goala.pdf"


def count_pages(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        num_pages = len(reader.pages)
    return num_pages


def convert_to_pdf(doc):
    word = win32.DispatchEx("Word.Application")
    new_name = doc.replace(".docx", r".pdf")
    worddoc = word.Documents.Open(doc)
    worddoc.SaveAs(new_name, FileFormat=17)
    worddoc.Close()
    return new_name


def get_today_date():
    today = dt.today().date()
    return today.strftime("%d-%m-%Y")


def get_db_connection():
    con_string = r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=G:/Shared drives/Root/11. DATABASE/DBRGT-02.accdb;"
    return pyodbc.connect(con_string)


def fetch_single_value(cursor, query, params):
    return cursor.execute(query, params).fetchval()


def get_Firma_proiectare(cursor, id_firma_proiectare):
    return {
        'nume': fetch_single_value(cursor, 'SELECT NumeFirma FROM tblFirmeProiectare WHERE IDFirma = ?', (id_firma_proiectare,)),
        'localitate': fetch_single_value(cursor, 'SELECT Localitate FROM tblFirmeProiectare WHERE IDFirma = ?', (id_firma_proiectare,)),
        'adresa': fetch_single_value(cursor, 'SELECT Adresa FROM tblFirmeProiectare WHERE IDFirma = ?', (id_firma_proiectare,)),
        'judet': fetch_single_value(cursor, 'SELECT Judet FROM tblFirmeProiectare WHERE IDFirma = ?', (id_firma_proiectare,)),
        'CUI': fetch_single_value(cursor, 'SELECT CUI FROM tblFirmeProiectare WHERE IDFirma = ?', (id_firma_proiectare,)),
        'NrRegCom': fetch_single_value(cursor, 'SELECT NrRegCom FROM tblFirmeProiectare WHERE IDFirma = ?', (id_firma_proiectare,)),
        'telefon': fetch_single_value(cursor, 'SELECT Telefon FROM tblFirmeProiectare WHERE IDFirma = ?', (id_firma_proiectare,)),
        'email': fetch_single_value(cursor, 'SELECT Email FROM tblFirmeProiectare WHERE IDFirma = ?', (id_firma_proiectare,)),
        'reprezentant': fetch_single_value(cursor, 'SELECT Nume FROM tblAngajati WHERE IDAngajat = (SELECT IDReprezentant FROM tblFirmeProiectare WHERE IDFirma = ?)', (id_firma_proiectare,)),
        'localitate_repr': fetch_single_value(cursor, 'SELECT Localitate FROM tblAngajati WHERE IDAngajat = (SELECT IDReprezentant FROM tblFirmeProiectare WHERE IDFirma = ?)', (id_firma_proiectare,)),
        'adresa_repr': fetch_single_value(cursor, 'SELECT Adresa FROM tblAngajati WHERE IDAngajat = (SELECT IDReprezentant FROM tblFirmeProiectare WHERE IDFirma = ?)', (id_firma_proiectare,)),
        'judet_repr': fetch_single_value(cursor, 'SELECT Judet FROM tblAngajati WHERE IDAngajat = (SELECT IDReprezentant FROM tblFirmeProiectare WHERE IDFirma = ?)', (id_firma_proiectare,)),
        'seria_CI': fetch_single_value(cursor, 'SELECT SerieCI FROM tblAngajati WHERE IDAngajat = (SELECT IDReprezentant FROM tblFirmeProiectare WHERE IDFirma = ?)', (id_firma_proiectare,)),
        'nr_CI': fetch_single_value(cursor, 'SELECT NumarCI FROM tblAngajati WHERE IDAngajat = (SELECT IDReprezentant FROM tblFirmeProiectare WHERE IDFirma = ?)', (id_firma_proiectare,)),
        'data_CI': fetch_single_value(cursor, 'SELECT DataCI FROM tblAngajati WHERE IDAngajat = (SELECT IDReprezentant FROM tblFirmeProiectare WHERE IDFirma = ?)', (id_firma_proiectare,)),
        'cnp_repr': fetch_single_value(cursor, 'SELECT CNP FROM tblAngajati WHERE IDAngajat = (SELECT IDReprezentant FROM tblFirmeProiectare WHERE IDFirma = ?)', (id_firma_proiectare,)),
        'caleCI': fetch_single_value(cursor, 'SELECT CaleCI FROM tblAngajati WHERE IDAngajat = (SELECT IDReprezentant FROM tblFirmeProiectare WHERE IDFirma = ?)', (id_firma_proiectare,)),
        'CaleStampila': fetch_single_value(cursor, 'SELECT CaleStampila FROM tblFirmeProiectare WHERE IDFirma = ?', (id_firma_proiectare,)),
        'CaleCertificat': fetch_single_value(cursor, 'SELECT CaleCertificat FROM tblFirmeProiectare WHERE IDFirma = ?', (id_firma_proiectare,)),
    }


def get_Client(cursor, id_client):
    return {
        'nume': fetch_single_value(cursor, 'SELECT NumeClient FROM tblClienti WHERE IDClient = ?', (id_client,)),
        'localitate': fetch_single_value(cursor, 'SELECT Localitate FROM tblClienti WHERE IDClient = ?', (id_client,)),
        'adresa': fetch_single_value(cursor, 'SELECT Adresa FROM tblClienti WHERE IDClient = ?', (id_client,)),
        'judet': fetch_single_value(cursor, 'SELECT Judet FROM tblClienti WHERE IDClient = ?', (id_client,)),
        'tip_client': fetch_single_value(cursor, 'SELECT TipClient FROM tblClienti WHERE IDClient = ?', (id_client,)),
        'CUI': fetch_single_value(cursor, 'SELECT CUI FROM tblClienti WHERE IDClient = ?', (id_client,)),
        'NrRegCom': fetch_single_value(cursor, 'SELECT NrRegCom FROM tblClienti WHERE IDClient = ?', (id_client,)),
        'reprezentant': fetch_single_value(cursor, 'SELECT Reprezentant FROM tblClienti WHERE IDClient = ?', (id_client,)),
        'telefon': fetch_single_value(cursor, 'SELECT Telefon FROM tblClienti WHERE IDClient = ?', (id_client,)),
        'email': fetch_single_value(cursor, 'SELECT Email FROM tblClienti WHERE IDClient = ?', (id_client,)),
    }


def get_UAT(cursor, IDUAT):
    return {
        'denumire_institutie': fetch_single_value(cursor, 'SELECT DenumireInstitutie FROM tblUAT WHERE IDUAT = ?', (IDUAT,)),
        'localitate': fetch_single_value(cursor, 'SELECT LocalitateInstitutie FROM tblUAT WHERE IDUAT = ?', (IDUAT,)),
        'adresa': fetch_single_value(cursor, 'SELECT AdresaInstitutie FROM tblUAT WHERE IDUAT = ?', (IDUAT,)),
        'judet': fetch_single_value(cursor, 'SELECT JudetInstitutie FROM tblUAT WHERE IDUAT = ?', (IDUAT,)),
        'cod_postal': fetch_single_value(cursor, 'SELECT CodPostal FROM tblUAT WHERE IDUAT = ?', (IDUAT,)),
        'telefon': fetch_single_value(cursor, 'SELECT Telefon FROM tblUAT WHERE IDUAT = ?', (IDUAT,)),
        'email': fetch_single_value(cursor, 'SELECT AdresaEmail FROM tblUAT WHERE IDUAT = ?', (IDUAT,)),
    }


def get_Beneficiar(cursor, id_beneficiar):
    return {
        'nume': fetch_single_value(cursor, 'SELECT NumeBeneficiar FROM tblBeneficiari WHERE IDBeneficiar = ?', (id_beneficiar,)),
        'localitate': fetch_single_value(cursor, 'SELECT Localitate FROM tblBeneficiari WHERE IDBeneficiar = ?', (id_beneficiar,)),
        'adresa': fetch_single_value(cursor, 'SELECT Adresa FROM tblBeneficiari WHERE IDBeneficiar = ?', (id_beneficiar,)),
        'judet': fetch_single_value(cursor, 'SELECT Judet FROM tblBeneficiari WHERE IDBeneficiar = ?', (id_beneficiar,)),
        'CodPostal': fetch_single_value(cursor, 'SELECT CodPostal FROM tblBeneficiari WHERE IDBeneficiar = ?', (id_beneficiar,)),
        'CUI': fetch_single_value(cursor, 'SELECT CUI FROM tblBeneficiari WHERE IDBeneficiar = ?', (id_beneficiar,)),
        'NrRegCom': fetch_single_value(cursor, 'SELECT NrRegCom FROM tblBeneficiari WHERE IDBeneficiar = ?', (id_beneficiar,)),
        'reprezentant': fetch_single_value(cursor, 'SELECT Reprezentant FROM tblBeneficiari WHERE IDBeneficiar = ?', (id_beneficiar,)),
        'telefon': fetch_single_value(cursor, 'SELECT Telefon FROM tblBeneficiari WHERE IDBeneficiar = ?', (id_beneficiar,)),
        'email': fetch_single_value(cursor, 'SELECT Email FROM tblBeneficiari WHERE IDBeneficiar = ?', (id_beneficiar,)),
    }


def get_Lucrare(cursor, id_lucrare):
    return {
        'nume': fetch_single_value(cursor, 'SELECT DenumireLucrare FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'localitate': fetch_single_value(cursor, 'SELECT LocalitateLucrare FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'adresa': fetch_single_value(cursor, 'SELECT AdresaLucrare FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'judet': fetch_single_value(cursor, 'SELECT JudetLucrare FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'IDFirmaProiectare': fetch_single_value(cursor, 'SELECT IDFirmaProiectare FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'IDClient': fetch_single_value(cursor, 'SELECT IDClient FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'IDBeneficiar': fetch_single_value(cursor, 'SELECT IDBeneficiar FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'descrierea_proiectului': fetch_single_value(cursor, 'SELECT DescriereaProiectului FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'emitent_cu': fetch_single_value(cursor, 'SELECT EmitentCU FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'nr_cu': fetch_single_value(cursor, 'SELECT NumarCU FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'data_cu': fetch_single_value(cursor, 'SELECT DataCU FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'IDPersoanaContact': fetch_single_value(cursor, 'SELECT IDPersoanaContact FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'IDIntocmit': fetch_single_value(cursor, 'SELECT IDIntocmit FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'IDVerificat': fetch_single_value(cursor, 'SELECT IDVerificat FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'facturare': fetch_single_value(cursor, 'SELECT FacturarePeClient FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'CaleCU': fetch_single_value(cursor, 'SELECT CaleCU FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'CalePlanIncadrare': fetch_single_value(cursor, 'SELECT CalePlanIncadrare FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'CalePlanSituatie': fetch_single_value(cursor, 'SELECT CalePlanSituatie FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'CaleMemoriuTehnic': fetch_single_value(cursor, 'SELECT CaleMemoriuTehnic FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'CaleActeBeneficiar': fetch_single_value(cursor, 'SELECT CaleActeBeneficiar FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'CaleActeFacturare': fetch_single_value(cursor, 'SELECT CaleActeFacturare FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'CaleChitantaAPM': fetch_single_value(cursor, 'SELECT CaleChitantaAPM FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'CaleChitantaDSP': fetch_single_value(cursor, 'SELECT CaleChitantaDSP FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'SuprafataMP': fetch_single_value(cursor, 'SELECT SuprafataMP FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'LungimeTraseuMetri': fetch_single_value(cursor, 'SELECT LungimeTraseuMetri FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'CaleACCUConstructie': fetch_single_value(cursor, 'SELECT CaleACCUConstructie FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'CaleAvizGiS': fetch_single_value(cursor, 'SELECT CaleAvizGiS FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'CaleAvizCTEsauATR': fetch_single_value(cursor, 'SELECT CaleAvizCTEsauATR FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'CaleExtraseCF': fetch_single_value(cursor, 'SELECT CaleExtraseCF FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'CalePlanSituatiePDF': fetch_single_value(cursor, 'SELECT CalePlanSituatiePDF FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'CalePlanSituatieDWG': fetch_single_value(cursor, 'SELECT CalePlanSituatieDWG FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
        'CaleRidicareTopoDWG': fetch_single_value(cursor, 'SELECT CaleRidicareTopoDWG FROM tblLucrari WHERE IDLucrare = ?', (id_lucrare,)),
    }


def get_Executie(cursor, id_executie):
    return {
        'nume': fetch_single_value(cursor, 'SELECT DenumireLucrare FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'localitate': fetch_single_value(cursor, 'SELECT LocalitateLucrare FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'adresa': fetch_single_value(cursor, 'SELECT AdresaLucrare FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'judet': fetch_single_value(cursor, 'SELECT JudetLucrare FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'IDFirmaProiectare': fetch_single_value(cursor, 'SELECT IDFirmaProiectare FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'IDClient': fetch_single_value(cursor, 'SELECT IDClient FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'IDBeneficiar': fetch_single_value(cursor, 'SELECT IDBeneficiar FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'IDUAT': fetch_single_value(cursor, 'SELECT IDEmitentAC FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'IDRTE': fetch_single_value(cursor, 'SELECT IDRTE FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'IDDS': fetch_single_value(cursor, 'SELECT IDDS FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'IDResponsabil': fetch_single_value(cursor, 'SELECT IDResponsabil FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'numar_ac': fetch_single_value(cursor, 'SELECT NumarAC FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'data_ac': fetch_single_value(cursor, 'SELECT DataAC FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'valabilitate_ac': fetch_single_value(cursor, 'SELECT ValabilitateAC FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'data_incepere_executie': fetch_single_value(cursor, 'SELECT DataIncepereExecutie FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'valabilitate_executie': fetch_single_value(cursor, 'SELECT ValabilitateExecutie FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),

        'nr_cl': fetch_single_value(cursor, 'SELECT NrCL FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'nr_anunt_UAT': fetch_single_value(cursor, 'SELECT NrAnuntUAT FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'nr_decizie_personal': fetch_single_value(cursor, 'SELECT NrDeciziePersonal FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'grafic_executie': fetch_single_value(cursor, 'SELECT GraficExecutie FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'data_incepere_grafic': fetch_single_value(cursor, 'SELECT DataIncepereGrafic FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'data_finalizare_grafic': fetch_single_value(cursor, 'SELECT DataFinalizareGrafic FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'lucrari_domeniu_public': fetch_single_value(cursor, 'SELECT LucrariDomeniuPublic FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),


        'CaleACScanat': fetch_single_value(cursor, 'SELECT CaleACScanat FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'CalePlanIncadrareACScanat': fetch_single_value(cursor, 'SELECT CalePlanIncadrareACScanat FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'CalePlanSituatieACScanat': fetch_single_value(cursor, 'SELECT CalePlanSituatieACScanat FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'CaleMemoriuTehnicACScanat': fetch_single_value(cursor, 'SELECT CaleMemoriuTehnicACScanat FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),


        'CalePlanIncadrarePTH': fetch_single_value(cursor, 'SELECT CalePlanIncadrarePTH FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'CalePlanSituatiePTH': fetch_single_value(cursor, 'SELECT CalePlanSituatiePTH FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'CaleSchemaMonofilaraJT': fetch_single_value(cursor, 'SELECT CaleSchemaMonofilaraJT FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'CaleSchemaMonofilaraMT': fetch_single_value(cursor, 'SELECT CaleSchemaMonofilaraMT FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),

        'CaleInstruireColectiva': fetch_single_value(cursor, 'SELECT CaleInstruireColectiva FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'CaleAvizCTEsauATR': fetch_single_value(cursor, 'SELECT CaleAvizCTEsauATR FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'CaleContract1': fetch_single_value(cursor, 'SELECT CaleContract1 FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'CaleContract2': fetch_single_value(cursor, 'SELECT CaleContract2 FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),
        'CaleContract3': fetch_single_value(cursor, 'SELECT CaleContract3 FROM tblExecutie WHERE IDExecutie = ?', (id_executie,)),

    }


def get_Contact(cursor, id_contact):
    return {
        'nume': fetch_single_value(cursor, 'SELECT Nume FROM tblAngajati WHERE IDAngajat = ?', (id_contact,)),
        'telefon': fetch_single_value(cursor, 'SELECT Telefon FROM tblAngajati WHERE IDAngajat = ?', (id_contact,)),
    }


def create_document(model_path, context, final_destination, stampila_path=None):
    doc = DocxTemplate(model_path)
    if stampila_path:
        doc.replace_pic("Placeholder_1.png", stampila_path)
    doc.render(context)
    nume = os.path.basename(model_path).strip('.docx')
    path_doc = os.path.join(
        final_destination, f'{nume}.docx')
    doc.save(path_doc)
    cerere_pdf_path = convert_to_pdf(path_doc)
    if os.path.exists(path_doc):
        os.remove(path_doc)
    return cerere_pdf_path


def create_email(model_path, context, final_destination):
    doc = DocxTemplate(model_path)
    doc.render(context)
    path_doc = os.path.join(final_destination, 'Email.docx')
    doc.save(path_doc)


def merge_pdfs(pdf_list, output_path):
    merger = PdfMerger()
    for pdf in pdf_list:
        merger.append(pdf)
    merger.write(output_path)
    merger.close()


def merge_pdfs_print(pdf_list, output_path):
    merger = PdfMerger()
    for pdf in pdf_list:
        merger.append(pdf)
        x = count_pages(pdf)
        if x % 2 == 1:
            merger.append(pagina_goala)
    merger.write(output_path)
    merger.close()


def get_data(path_final, director_final, id_lucrare):
    final_destination = os.path.join(path_final, director_final)
    os.makedirs(final_destination, exist_ok=True,)
    conn = get_db_connection()
    cursor = conn.cursor()
    astazi = get_today_date()

    lucrare = get_Lucrare(cursor, id_lucrare)
    firma_proiectare = get_Firma_proiectare(
        cursor, lucrare['IDFirmaProiectare'])
    client = get_Client(cursor, lucrare['IDClient'])
    beneficiar = get_Beneficiar(cursor, lucrare['IDBeneficiar'])
    contact = get_Contact(cursor, lucrare['IDPersoanaContact'])
    intocmit = fetch_single_value(
        cursor, 'SELECT Nume FROM tblAngajati WHERE IDAngajat = ?', (lucrare['IDIntocmit'],))
    verificat = fetch_single_value(
        cursor, 'SELECT Nume FROM tblAngajati WHERE IDAngajat = ?', (lucrare['IDVerificat'],))

    return {
        'astazi': astazi,
        'lucrare': lucrare,
        'firma_proiectare': firma_proiectare,
        'client': client,
        'beneficiar': beneficiar,
        'contact': contact,
        'final_destination': final_destination,
        'intocmit': intocmit,
        'verificat': verificat,
    }


def get_data_executie(path_final, director_final, id_executie):
    final_destination = os.path.join(path_final, director_final)
    os.makedirs(final_destination, exist_ok=True,)
    conn = get_db_connection()
    cursor = conn.cursor()
    astazi = get_today_date()

    executie = get_Executie(cursor, id_executie)
    uat = get_UAT(cursor, executie['IDUAT'])
    rte = get_UAT(cursor, executie['IDRTE'])
    diriginte_santier = get_UAT(cursor, executie['IDDS'])
    responsabil_constructii = get_UAT(cursor, executie['IDResponsabil'])
    firma_proiectare = get_Firma_proiectare(
        cursor, executie['IDFirmaProiectare'])
    client = get_Client(cursor, executie['IDClient'])
    beneficiar = get_Beneficiar(cursor, executie['IDBeneficiar'])

    return {
        'astazi': astazi,
        'firma_proiectare': firma_proiectare,
        'client': client,
        'beneficiar': beneficiar,
        'uat': uat,
        'rte': rte,
        'diriginte_santier': diriginte_santier,
        'responsabil_contructii': responsabil_constructii,
        'final_destination': final_destination,
        'executie': executie,
    }


def facturare(id_lucrare):
    conn = get_db_connection()
    cursor = conn.cursor()
    lucrare = get_Lucrare(cursor, id_lucrare)
    firma_proiectare = get_Firma_proiectare(
        cursor, lucrare['IDFirmaProiectare'])
    client = get_Client(cursor, lucrare['IDClient'])
    if lucrare['facturare'] == False:
        return {
            'firma_facturare': firma_proiectare['nume'],
            'cui_firma_facturare': firma_proiectare['CUI'],
        }
    else:
        return {
            'firma_facturare': client['nume'],
            'cui_firma_facturare': client['CUI'],
        }


def copy_file(file_path, path_final, director_final, file_name: str):
    file = file_path.strip('"')
    shutil.copy(file, os.path.join(path_final, director_final, file_name))


def move_file(file_path, path_final, director_final, file_name: str):
    file = file_path.strip('"')
    shutil.move(file, os.path.join(path_final, director_final, file_name))


def copy_file_prefix(file_path, path_final, director_final, prefix=None):
    file = file_path.strip('"')
    filename = os.path.basename(file_path)
    new_filename = prefix + filename
    shutil.copy(file, os.path.join(path_final, director_final, new_filename))


def count_pages_ISU(cerere_path, cu_path, plan_incadrare_path, plan_situatie_path, cale_memoriu, cale_acte):
    cerere = count_pages(cerere_path)
    cu = count_pages(cu_path)
    plan_incadrare = count_pages(plan_incadrare_path)
    plan_situatie = count_pages(plan_situatie_path)
    memoriu = count_pages(cale_memoriu)
    acte = count_pages(cale_acte)
    return {
        'cerere': cerere,
        'cu': cu,
        'plan_incadrare': plan_incadrare,
        'plan_situatie': plan_situatie,
        'memoriu_tehnic': memoriu,
        'acte_facturare': acte
    }
