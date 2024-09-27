import os
import functii as x


def aviz_APM(id_lucrare, path_final):
    director_final = '01.Mediu APM Iasi'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere_APM = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/IS/01.Mediu APM/'f"01.Cerere{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

    context_Cerere_APM = {
        'nume_firma_proiectare': y['firma_proiectare']['nume'],
        'nume_beneficiar': y['beneficiar']['nume'],
        'nume_client': y['client']['nume'],
        'localitate_firma_proiectare': y['firma_proiectare']['localitate'],
        'adresa_firma_proiectare': y['firma_proiectare']['adresa'],
        'judet_firma_proiectare': y['firma_proiectare']['judet'],
        'email_firma_proiectare': y['firma_proiectare']['email'],
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
        'data': y['astazi'],
    }

    cerere_pdf_path = x.create_document(
        model_cerere_APM, context_Cerere_APM, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    # -----------------------------------------------------------------------------------------------

    # creez NOTIFICAREA

    model_Notificare_APM = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\IS\01.Mediu APM\02.Notificare.docx")

    context_Notificare_APM = {
        # lucrare
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        # client
        'nume_client': y['client']['nume'],
        'localitate_client': y['client']['localitate'],
        'adresa_client': y['client']['adresa'],
        'judet_client': y['client']['judet'],
        # beneficiar
        'nume_beneficiar': y['beneficiar']['nume'],
        # proiectare
        'nume_firma_proiectare': y['firma_proiectare']['nume'],
        'localitate_firma_proiectare': y['firma_proiectare']['localitate'],
        'adresa_firma_proiectare': y['firma_proiectare']['adresa'],
        'judet_firma_proiectare': y['firma_proiectare']['judet'],
        'email_firma_proiectare': y['firma_proiectare']['email'],
        'reprezentant_firma_proiectare': y['firma_proiectare']['reprezentant'],
        'descrierea_proiectului': y['lucrare']['descrierea_proiectului'],
        'intocmit': y['intocmit'],
        'verificat': y['verificat'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
    }

    notificare_pdf_path = x.create_document(
        model_Notificare_APM, context_Notificare_APM, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    # -----------------------------------------------------------------------------------------------

    # creez EMAILUL

    model_Email_APM = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\IS\01.Mediu APM\Model email.docx")

    context_Email_APM_Iasi = {
        'nume_client': y['client']['nume'],
        'nr_cu': y['lucrare']['nr_cu'],
        'data_cu': y['lucrare']['data_cu'],
        'emitent_cu': y['lucrare']['emitent_cu'],
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nume_client': y['client']['nume'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
    }

    x.create_email(model_Email_APM, context_Email_APM_Iasi,
                   y['final_destination'])

    # -----------------------------------------------------------------------------------------------

    # creez DOCUMENTUL FINAL

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz APM Iași - {y['client']['nume']} conform CU {y['lucrare']['nr_cu']} din {y['lucrare']['data_cu']}.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['lucrare']['CaleChitantaAPM'].strip('"'),
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        notificare_pdf_path,
        y['lucrare']['CaleActeBeneficiar'].strip('"'),
    ]

    x.merge_pdfs(pdf_list, path_document_final)

    # -----------------------------------------------------------------------------------------------

    # fac curat

    if os.path.exists(cerere_pdf_path):
        os.remove(cerere_pdf_path)

    if os.path.exists(notificare_pdf_path):
        os.remove(notificare_pdf_path)

    print("\nAvizul APM Iași a fost creat \n")


def aviz_EE_Delgaz(id_lucrare, path_final):
    director_final = '02.Aviz EE Delgaz'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere_EE_Delgaz = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/IS/02.Aviz EE Delgaz - Iasi/'f"01.Cerere aviz EE Delgaz{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

    context_cerere_EE_Delgaz = {
        # proiectare

        'nume_firma_proiectare': y['firma_proiectare']['nume'],
        'localitate_firma_proiectare': y['firma_proiectare']['localitate'],
        'adresa_firma_proiectare': y['firma_proiectare']['adresa'],
        'judet_firma_proiectare': y['firma_proiectare']['judet'],
        'email_firma_proiectare': y['firma_proiectare']['email'],
        'cui_firma_proiectare': y['firma_proiectare']['CUI'],
        'nr_reg_com': y['firma_proiectare']['NrRegCom'],
        'reprezentant_firma_proiectare': y['firma_proiectare']['reprezentant'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
        # client
        'nume_client': y['client']['nume'],
        'localitate_client': y['client']['localitate'],
        'adresa_client': y['client']['adresa'],
        'judet_client': y['client']['judet'],
        # beneficiar
        'nume_beneficiar': y['beneficiar']['nume'],
        # lucrare
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nr_cu': y['lucrare']['nr_cu'],
        'data_cu': y['lucrare']['data_cu'],
        'emitent_cu': y['lucrare']['emitent_cu'],
        # Data
        'data': y['astazi'],
    }

    cerere_EE_pdf_path = x.create_document(
        model_cerere_EE_Delgaz, context_cerere_EE_Delgaz, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz EE - Delgaz - {y['client']['nume']} conform CU {y['lucrare']['nr_cu']} din {y['lucrare']['data_cu']}.pdf")

    pdf_list = [
        cerere_EE_pdf_path,
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CaleMemoriuTehnic'].strip('"'),
        y['lucrare']['CaleActeFacturare'].strip('"'),
    ]

    x.merge_pdfs(pdf_list, path_document_final)

    if os.path.exists(cerere_EE_pdf_path):
        os.remove(cerere_EE_pdf_path)

    # -----------------------------------------------------------------------------------------------

    # creez EMAILUL

    model_email = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/IS/02.Aviz EE Delgaz - Iasi/'f"Model email{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

    facturare = x.facturare(id_lucrare)

    context_email = {
        'nume_client': y['client']['nume'],
        'nr_cu': y['lucrare']['nr_cu'],
        'data_cu': y['lucrare']['data_cu'],
        'emitent_cu': y['lucrare']['emitent_cu'],
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nume_client': y['client']['nume'],
        'firma_facturare': facturare['firma_facturare'],
        'cui_firma_facturare': facturare['cui_firma_facturare'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
    }

    x.create_email(model_email, context_email, y['final_destination'])

    print("\nAvizul EE Delgaz a fost creat \n")


def aviz_GN_Delgaz(id_lucrare, path_final):
    director_final = '03.Aviz GN Delgaz'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/IS/03.Aviz GN Delgaz/'f"01.Aviz GN{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

    context_cerere = {
        # proiectare

        'nume_firma_proiectare': y['firma_proiectare']['nume'],
        'localitate_firma_proiectare': y['firma_proiectare']['localitate'],
        'adresa_firma_proiectare': y['firma_proiectare']['adresa'],
        'judet_firma_proiectare': y['firma_proiectare']['judet'],
        'email_firma_proiectare': y['firma_proiectare']['email'],
        'cui_firma_proiectare': y['firma_proiectare']['CUI'],
        'nr_reg_com': y['firma_proiectare']['NrRegCom'],
        'reprezentant_firma_proiectare': y['firma_proiectare']['reprezentant'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
        # client
        'nume_client': y['client']['nume'],
        'localitate_client': y['client']['localitate'],
        'adresa_client': y['client']['adresa'],
        'judet_client': y['client']['judet'],
        # beneficiar
        'nume_beneficiar': y['beneficiar']['nume'],
        # lucrare
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nr_cu': y['lucrare']['nr_cu'],
        'data_cu': y['lucrare']['data_cu'],
        'emitent_cu': y['lucrare']['emitent_cu'],
        # Data
        'data': y['astazi'],
    }

    cerere_GN_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz GN - Delgaz - {y['client']['nume']} conform CU nr. {y['lucrare']['nr_cu']} din {y['lucrare']['data_cu']}.pdf")

    pdf_list = [
        cerere_GN_pdf_path,
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CaleMemoriuTehnic'].strip('"'),
        y['lucrare']['CaleActeFacturare'].strip('"'),
    ]

    x.merge_pdfs(pdf_list, path_document_final)

    if os.path.exists(cerere_GN_pdf_path):
        os.remove(cerere_GN_pdf_path)

    # -----------------------------------------------------------------------------------------------

    # creez EMAILUL

    model_email = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/IS/03.Aviz GN Delgaz/'f"Model email{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

    facturare = x.facturare(id_lucrare)

    context_email = {
        'nume_client': y['client']['nume'],
        'nr_cu': y['lucrare']['nr_cu'],
        'data_cu': y['lucrare']['data_cu'],
        'emitent_cu': y['lucrare']['emitent_cu'],
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nume_client': y['client']['nume'],
        'firma_facturare': facturare['firma_facturare'],
        'cui_firma_facturare': facturare['cui_firma_facturare'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
    }

    x.create_email(model_email, context_email, y['final_destination'])

    print("\nAvizul GN Delgaz a fost creat \n")


def aviz_Gazmir(id_lucrare, path_final):
    director_final = '04.Aviz GN Gazmir'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/IS/04.Aviz GN Gazmir/'f"01.Model aviz Gazmir{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

    context_cerere = {
        # proiectare

        'nume_firma_proiectare': y['firma_proiectare']['nume'],
        'localitate_firma_proiectare': y['firma_proiectare']['localitate'],
        'adresa_firma_proiectare': y['firma_proiectare']['adresa'],
        'judet_firma_proiectare': y['firma_proiectare']['judet'],
        'email_firma_proiectare': y['firma_proiectare']['email'],
        'cui_firma_proiectare': y['firma_proiectare']['CUI'],
        'nr_reg_com': y['firma_proiectare']['NrRegCom'],
        'reprezentant_firma_proiectare': y['firma_proiectare']['reprezentant'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
        # client
        'nume_client': y['client']['nume'],
        'localitate_client': y['client']['localitate'],
        'adresa_client': y['client']['adresa'],
        'judet_client': y['client']['judet'],
        # beneficiar
        'nume_beneficiar': y['beneficiar']['nume'],
        # lucrare
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nr_cu': y['lucrare']['nr_cu'],
        'data_cu': y['lucrare']['data_cu'],
        'emitent_cu': y['lucrare']['emitent_cu'],
        # Data
        'data': y['astazi'],
    }

    cerere_Gazmir_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz GN Gazmir - {y['client']['nume']} conform CU nr. {y['lucrare']['nr_cu']} din {y['lucrare']['data_cu']}.pdf")

    pdf_list = [
        cerere_Gazmir_pdf_path,
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CaleMemoriuTehnic'].strip('"'),
        y['lucrare']['CaleActeFacturare'].strip('"'),
    ]

    x.merge_pdfs(pdf_list, path_document_final)

    if os.path.exists(cerere_Gazmir_pdf_path):
        os.remove(cerere_Gazmir_pdf_path)

    # -----------------------------------------------------------------------------------------------

    # creez EMAILUL

    model_email = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\IS\04.Aviz GN Gazmir\Model email.docx")

    facturare = x.facturare(id_lucrare)

    context_email = {
        'nume_client': y['client']['nume'],
        'nr_cu': y['lucrare']['nr_cu'],
        'data_cu': y['lucrare']['data_cu'],
        'emitent_cu': y['lucrare']['emitent_cu'],
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nume_client': y['client']['nume'],
        'firma_facturare': facturare['firma_facturare'],
        'cui_firma_facturare': facturare['cui_firma_facturare'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],

    }

    x.create_email(model_email, context_email, y['final_destination'])

    print("\nAvizul GN Gazmir a fost creat \n")


def aviz_Apavital(id_lucrare, path_final):
    director_final = '05.Aviz Apavital'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\IS\05.Aviz Apavital\01.Cerere.docx")

    facturare = x.facturare(id_lucrare)

    context_cerere = {
        # proiectare
        'nume_firma_proiectare': y['firma_proiectare']['nume'],
        'localitate_firma_proiectare': y['firma_proiectare']['localitate'],
        'adresa_firma_proiectare': y['firma_proiectare']['adresa'],
        'judet_firma_proiectare': y['firma_proiectare']['judet'],
        'email_firma_proiectare': y['firma_proiectare']['email'],
        'cui_firma_proiectare': y['firma_proiectare']['CUI'],
        'nr_reg_com': y['firma_proiectare']['NrRegCom'],
        'reprezentant_firma_proiectare': y['firma_proiectare']['reprezentant'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
        # client
        'nume_client': y['client']['nume'],
        'localitate_client': y['client']['localitate'],
        'adresa_client': y['client']['adresa'],
        'judet_client': y['client']['judet'],
        # beneficiar
        'nume_beneficiar': y['beneficiar']['nume'],
        'localitate_beneficiar': y['beneficiar']['localitate'],
        'adresa_beneficiar': y['beneficiar']['adresa'],
        'judet_beneficiar': y['beneficiar']['judet'],
        # lucrare
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nr_cu': y['lucrare']['nr_cu'],
        'data_cu': y['lucrare']['data_cu'],
        'emitent_cu': y['lucrare']['emitent_cu'],
        # FACTURARE
        'firma_facturare': facturare['firma_facturare'],
        'cui_firma_facturare': facturare['cui_firma_facturare'],
        # Data
        'data': y['astazi'],
    }

    cerere_Gazmir_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))
    os.rename(cerere_Gazmir_pdf_path, os.path.join(
        path_final, director_final, '01.Cerere.pdf'))

    x.copy_file(y['lucrare']['CaleCU'], path_final,
                director_final, '02.Certificat de urbanism.pdf')
    x.copy_file(y['lucrare']['CalePlanIncadrare'], path_final,
                director_final, '03.Plan incadrare in zona.pdf')
    x.copy_file(y['lucrare']['CalePlanSituatie'], path_final,
                director_final, '04.Plan situatie.pdf')
    x.copy_file(y['lucrare']['CaleMemoriuTehnic'], path_final,
                director_final, '05.Memoriu tehnic.pdf')
    x.copy_file(y['lucrare']['CaleActeFacturare'], path_final,
                director_final, '06.Acte facturare.pdf')
    # -----------------------------------------------------------------------------------------------

    print("\nAvizul Apavital a fost creat \n")


def aviz_Termoficare(id_lucrare, path_final):
    director_final = '06.Aviz Termoficare'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/IS/06.Aviz Termoficare/'f"01.Model aviz Termoficare{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

    context_cerere = {
        # proiectare

        'nume_firma_proiectare': y['firma_proiectare']['nume'],
        'localitate_firma_proiectare': y['firma_proiectare']['localitate'],
        'adresa_firma_proiectare': y['firma_proiectare']['adresa'],
        'judet_firma_proiectare': y['firma_proiectare']['judet'],
        'email_firma_proiectare': y['firma_proiectare']['email'],
        'cui_firma_proiectare': y['firma_proiectare']['CUI'],
        'nr_reg_com': y['firma_proiectare']['NrRegCom'],
        'reprezentant_firma_proiectare': y['firma_proiectare']['reprezentant'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
        # client
        'nume_client': y['client']['nume'],
        'localitate_client': y['client']['localitate'],
        'adresa_client': y['client']['adresa'],
        'judet_client': y['client']['judet'],
        # beneficiar
        'nume_beneficiar': y['beneficiar']['nume'],
        # lucrare
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nr_cu': y['lucrare']['nr_cu'],
        'data_cu': y['lucrare']['data_cu'],
        'emitent_cu': y['lucrare']['emitent_cu'],
        # Data
        'data': y['astazi'],
    }

    cerere_Termoficare_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz Termoficare - DE PRINTAT.pdf")

    pdf_list = [
        cerere_Termoficare_pdf_path,
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CaleMemoriuTehnic'].strip('"'),
        y['lucrare']['CaleActeFacturare'].strip('"'),
    ]

    x.merge_pdfs_print(pdf_list, path_document_final)

    if os.path.exists(cerere_Termoficare_pdf_path):
        os.remove(cerere_Termoficare_pdf_path)

    x.copy_file("G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/IS/06.Aviz Termoficare/Citeste-ma.docx",
                path_final, director_final, 'Citeste-ma.docx')

    print("\nAvizul Termoficare a fost creat \n")


def aviz_Orange(id_lucrare, path_final):
    director_final = '07.Aviz Orange'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/IS/07.Aviz Orange/'f"Cerere Orange{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

    context_cerere = {
        # proiectare
        'nume_firma_proiectare': y['firma_proiectare']['nume'],
        'localitate_firma_proiectare': y['firma_proiectare']['localitate'],
        'adresa_firma_proiectare': y['firma_proiectare']['adresa'],
        'judet_firma_proiectare': y['firma_proiectare']['judet'],
        'email_firma_proiectare': y['firma_proiectare']['email'],
        'cui_firma_proiectare': y['firma_proiectare']['CUI'],
        'nr_reg_com': y['firma_proiectare']['NrRegCom'],
        'reprezentant_firma_proiectare': y['firma_proiectare']['reprezentant'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
        # client
        'nume_client': y['client']['nume'],
        'localitate_client': y['client']['localitate'],
        'adresa_client': y['client']['adresa'],
        'judet_client': y['client']['judet'],
        # beneficiar
        'nume_beneficiar': y['beneficiar']['nume'],
        'localitate_beneficiar': y['beneficiar']['localitate'],
        'adresa_beneficiar': y['beneficiar']['adresa'],
        'judet_beneficiar': y['beneficiar']['judet'],
        # lucrare
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nr_cu': y['lucrare']['nr_cu'],
        'data_cu': y['lucrare']['data_cu'],
        'emitent_cu': y['lucrare']['emitent_cu'],
        # Data
        'data': y['astazi'],
    }

    cerere_Orange_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))
    os.rename(cerere_Orange_pdf_path, os.path.join(
        path_final, director_final, '01.Cerere.pdf'))

    x.copy_file(y['lucrare']['CaleCU'], path_final,
                director_final, '02.Certificat de urbanism.pdf')
    x.copy_file(y['lucrare']['CalePlanIncadrare'], path_final,
                director_final, '03.Plan incadrare in zona.pdf')
    x.copy_file(y['lucrare']['CalePlanSituatie'], path_final,
                director_final, '04.Plan situatie.pdf')
    x.copy_file(y['lucrare']['CaleMemoriuTehnic'], path_final,
                director_final, '05.Memoriu tehnic.pdf')
    x.copy_file(y['lucrare']['CaleActeFacturare'], path_final,
                director_final, '06.Acte facturare.pdf')
    x.copy_file("G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/IS/07.Aviz Orange/Citeste-ma.docx",
                path_final, director_final, 'Citeste-ma.docx')
    # -----------------------------------------------------------------------------------------------

    print("\nAvizul Orange a fost creat \n")


def aviz_CTP(id_lucrare, path_final):
    director_final = '08.Aviz CTP'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/IS/08.Aviz CTP/'f"01.Model aviz CTP{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

    context_cerere = {
        # proiectare
        'nume_firma_proiectare': y['firma_proiectare']['nume'],
        'localitate_firma_proiectare': y['firma_proiectare']['localitate'],
        'adresa_firma_proiectare': y['firma_proiectare']['adresa'],
        'judet_firma_proiectare': y['firma_proiectare']['judet'],
        'email_firma_proiectare': y['firma_proiectare']['email'],
        'cui_firma_proiectare': y['firma_proiectare']['CUI'],
        'nr_reg_com': y['firma_proiectare']['NrRegCom'],
        'reprezentant_firma_proiectare': y['firma_proiectare']['reprezentant'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
        # client
        'nume_client': y['client']['nume'],
        'localitate_client': y['client']['localitate'],
        'adresa_client': y['client']['adresa'],
        'judet_client': y['client']['judet'],
        # beneficiar
        'nume_beneficiar': y['beneficiar']['nume'],
        # lucrare
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nr_cu': y['lucrare']['nr_cu'],
        'data_cu': y['lucrare']['data_cu'],
        'emitent_cu': y['lucrare']['emitent_cu'],
        # Data
        'data': y['astazi'],
    }

    cerere_CTP_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(path_final, director_final, f"Documentatie aviz CTP - DE PRINTAT.pdf")

    pdf_list = [
        cerere_CTP_pdf_path,
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CaleMemoriuTehnic'].strip('"'),
        y['lucrare']['CaleActeFacturare'].strip('"'),
    ]

    x.merge_pdfs_print(pdf_list, path_document_final)

    if os.path.exists(cerere_CTP_pdf_path):
        os.remove(cerere_CTP_pdf_path)

    print("\nAvizul CTP a fost creat \n")


def aviz_PMI_Mediu(id_lucrare, path_final):
    director_final = '10.Aviz PMI - Mediu'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\IS\10.Aviz PMI Mediu\01.Cerere.docx")

    context_cerere = {
        'nume_firma_proiectare': y['firma_proiectare']['nume'],
        'nume_beneficiar': y['beneficiar']['nume'],
        'nume_client': y['client']['nume'],
        'localitate_firma_proiectare': y['firma_proiectare']['localitate'],
        'adresa_firma_proiectare': y['firma_proiectare']['adresa'],
        'judet_firma_proiectare': y['firma_proiectare']['judet'],
        'email_firma_proiectare': y['firma_proiectare']['email'],
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
        'data': y['astazi'],
    }

    cerere_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    # -----------------------------------------------------------------------------------------------

    # creez PLAN MEDIU

    model_Plan_Mediu = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\IS\10.Aviz PMI Mediu\02.Plan Mediu PMI.docx")

    context_Plan_Mediu = {
        # lucrare
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        # client
        'nume_client': y['client']['nume'],
        'localitate_client': y['client']['localitate'],
        'adresa_client': y['client']['adresa'],
        'judet_client': y['client']['judet'],
        # beneficiar
        'nume_beneficiar': y['beneficiar']['nume'],
        # proiectare
        'nume_firma_proiectare': y['firma_proiectare']['nume'],
        'localitate_firma_proiectare': y['firma_proiectare']['localitate'],
        'adresa_firma_proiectare': y['firma_proiectare']['adresa'],
        'judet_firma_proiectare': y['firma_proiectare']['judet'],
        'email_firma_proiectare': y['firma_proiectare']['email'],
        'reprezentant_firma_proiectare': y['firma_proiectare']['reprezentant'],
        'descrierea_proiectului': y['lucrare']['descrierea_proiectului'],
        'intocmit': y['intocmit'],
        'verificat': y['verificat'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
    }

    plan_mediu_pdf_path = x.create_document(
        model_Plan_Mediu, context_Plan_Mediu, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    # -----------------------------------------------------------------------------------------------

    # creez EMAILUL

    model_email = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\IS\10.Aviz PMI Mediu\Model email.docx")

    context_email = {
        'nume_client': y['client']['nume'],
        'nr_cu': y['lucrare']['nr_cu'],
        'data_cu': y['lucrare']['data_cu'],
        'emitent_cu': y['lucrare']['emitent_cu'],
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nume_client': y['client']['nume'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],

    }

    x.create_email(model_email, context_email, y['final_destination'])

    # -----------------------------------------------------------------------------------------------

    # creez DOCUMENTUL FINAL

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz PMI Mediu - {y['client']['nume']} conform CU {y['lucrare']['nr_cu']} din {y['lucrare']['data_cu']}.pdf")

    pdf_list = [
        cerere_pdf_path,
        plan_mediu_pdf_path,
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CaleMemoriuTehnic'].strip('"'),
        y['lucrare']['CaleActeBeneficiar'].strip('"'),
    ]

    x.merge_pdfs_print(pdf_list, path_document_final)

    # -----------------------------------------------------------------------------------------------

    # fac curat

    if os.path.exists(cerere_pdf_path):
        os.remove(cerere_pdf_path)

    if os.path.exists(plan_mediu_pdf_path):
        os.remove(plan_mediu_pdf_path)

    print("\nAvizul PMI - Mediu a fost creat \n")


def aviz_PMI_SEn(id_lucrare, path_final):
    director_final = '11.Aviz PMI SEn'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/IS/11.Aviz PMI SEn/'f"Model aviz PMI SEn{' - GENERAL TEHNIC' if y['lucrare']['IDFirmaProiectare'] == 3 else ' - PROING SERV' if y['lucrare']['IDFirmaProiectare'] == 4 else ' - ROGOTEHNIC'}.docx")

    context_cerere = {
        # proiectare

        'nume_firma_proiectare': y['firma_proiectare']['nume'],
        'localitate_firma_proiectare': y['firma_proiectare']['localitate'],
        'adresa_firma_proiectare': y['firma_proiectare']['adresa'],
        'judet_firma_proiectare': y['firma_proiectare']['judet'],
        'email_firma_proiectare': y['firma_proiectare']['email'],
        'cui_firma_proiectare': y['firma_proiectare']['CUI'],
        'nr_reg_com': y['firma_proiectare']['NrRegCom'],
        'reprezentant_firma_proiectare': y['firma_proiectare']['reprezentant'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
        # client
        'nume_client': y['client']['nume'],
        'localitate_client': y['client']['localitate'],
        'adresa_client': y['client']['adresa'],
        'judet_client': y['client']['judet'],
        # beneficiar
        'nume_beneficiar': y['beneficiar']['nume'],
        # lucrare
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nr_cu': y['lucrare']['nr_cu'],
        'data_cu': y['lucrare']['data_cu'],
        'emitent_cu': y['lucrare']['emitent_cu'],
        # Data
        'data': y['astazi'],
    }

    cerere_PMI_SEn_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz PMI - Serviciul energetic  - {y['client']['nume']} conform CU {y['lucrare']['nr_cu']} din {y['lucrare']['data_cu']}.pdf")

    pdf_list = [
        cerere_PMI_SEn_pdf_path,
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CaleMemoriuTehnic'].strip('"'),
        y['lucrare']['CaleActeBeneficiar'].strip('"'),
    ]

    x.merge_pdfs(pdf_list, path_document_final)

    if os.path.exists(cerere_PMI_SEn_pdf_path):
        os.remove(cerere_PMI_SEn_pdf_path)

    # -----------------------------------------------------------------------------------------------

    # creez EMAILUL

    model_email = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\IS\11.Aviz PMI SEn\Model email.docx")

    context_email = {
        'nume_client': y['client']['nume'],
        'nr_cu': y['lucrare']['nr_cu'],
        'data_cu': y['lucrare']['data_cu'],
        'emitent_cu': y['lucrare']['emitent_cu'],
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nume_client': y['client']['nume'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
    }

    x.create_email(model_email, context_email, y['final_destination'])

    print("\nAvizul PMI - SEn a fost creat \n")


def aviz_PMI_BSM(id_lucrare, path_final):
    director_final = '12.Aviz PMI Strazi Municipale'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/IS/12.Aviz PMI Strazi Municipale/'f"Model aviz PMI BSM{' - GENERAL TEHNIC' if y['lucrare']['IDFirmaProiectare'] == 3 else ' - PROING SERV' if y['lucrare']['IDFirmaProiectare'] == 4 else ' - ROGOTEHNIC'}.docx")

    context_cerere = {
        # proiectare

        'nume_firma_proiectare': y['firma_proiectare']['nume'],
        'localitate_firma_proiectare': y['firma_proiectare']['localitate'],
        'adresa_firma_proiectare': y['firma_proiectare']['adresa'],
        'judet_firma_proiectare': y['firma_proiectare']['judet'],
        'email_firma_proiectare': y['firma_proiectare']['email'],
        'cui_firma_proiectare': y['firma_proiectare']['CUI'],
        'nr_reg_com': y['firma_proiectare']['NrRegCom'],
        'reprezentant_firma_proiectare': y['firma_proiectare']['reprezentant'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
        # client
        'nume_client': y['client']['nume'],
        'localitate_client': y['client']['localitate'],
        'adresa_client': y['client']['adresa'],
        'judet_client': y['client']['judet'],
        # beneficiar
        'nume_beneficiar': y['beneficiar']['nume'],
        # lucrare
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nr_cu': y['lucrare']['nr_cu'],
        'data_cu': y['lucrare']['data_cu'],
        'emitent_cu': y['lucrare']['emitent_cu'],
        # Data
        'data': y['astazi'],
    }

    cerere_PMI_BSM_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz PMI - Strazi Municipale  - {y['client']['nume']} conform CU {y['lucrare']['nr_cu']} din {y['lucrare']['data_cu']}.pdf")

    pdf_list = [
        cerere_PMI_BSM_pdf_path,
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CaleMemoriuTehnic'].strip('"'),
        y['lucrare']['CaleActeBeneficiar'].strip('"'),
    ]

    x.merge_pdfs(pdf_list, path_document_final)

    if os.path.exists(cerere_PMI_BSM_pdf_path):
        os.remove(cerere_PMI_BSM_pdf_path)

    # -----------------------------------------------------------------------------------------------

    # creez EMAILUL

    model_email = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\IS\12.Aviz PMI Strazi Municipale\Model email.docx")

    context_email = {
        'nume_client': y['client']['nume'],
        'nr_cu': y['lucrare']['nr_cu'],
        'data_cu': y['lucrare']['data_cu'],
        'emitent_cu': y['lucrare']['emitent_cu'],
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nume_client': y['client']['nume'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
    }

    x.create_email(model_email, context_email, y['final_destination'])

    print("\nAvizul PMI - BSM a fost creat \n")


def aviz_PMI_Spatii_Verzi(id_lucrare, path_final):
    director_final = '13.Aviz PMI Spatii Verzi'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/IS/13.Aviz PMI Spatii Verzi/'f"Model aviz PMI Spatii Verzi{' - GENERAL TEHNIC' if y['lucrare']['IDFirmaProiectare'] == 3 else ' - PROING SERV' if y['lucrare']['IDFirmaProiectare'] == 4 else ' - ROGOTEHNIC'}.docx")

    context_cerere = {
        # proiectare

        'nume_firma_proiectare': y['firma_proiectare']['nume'],
        'localitate_firma_proiectare': y['firma_proiectare']['localitate'],
        'adresa_firma_proiectare': y['firma_proiectare']['adresa'],
        'judet_firma_proiectare': y['firma_proiectare']['judet'],
        'email_firma_proiectare': y['firma_proiectare']['email'],
        'cui_firma_proiectare': y['firma_proiectare']['CUI'],
        'nr_reg_com': y['firma_proiectare']['NrRegCom'],
        'reprezentant_firma_proiectare': y['firma_proiectare']['reprezentant'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
        # client
        'nume_client': y['client']['nume'],
        'localitate_client': y['client']['localitate'],
        'adresa_client': y['client']['adresa'],
        'judet_client': y['client']['judet'],
        # beneficiar
        'nume_beneficiar': y['beneficiar']['nume'],
        # lucrare
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nr_cu': y['lucrare']['nr_cu'],
        'data_cu': y['lucrare']['data_cu'],
        'emitent_cu': y['lucrare']['emitent_cu'],
        # Data
        'data': y['astazi'],
    }

    cerere_PMI_BSM_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz PMI - Spatii Verzi  - {y['client']['nume']} conform CU {y['lucrare']['nr_cu']} din {y['lucrare']['data_cu']}.pdf")

    pdf_list = [
        cerere_PMI_BSM_pdf_path,
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CaleMemoriuTehnic'].strip('"'),
        y['lucrare']['CaleActeBeneficiar'].strip('"'),
    ]

    x.merge_pdfs(pdf_list, path_document_final)

    if os.path.exists(cerere_PMI_BSM_pdf_path):
        os.remove(cerere_PMI_BSM_pdf_path)

    # -----------------------------------------------------------------------------------------------

    # creez EMAILUL

    model_email = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\IS\13.Aviz PMI Spatii Verzi\Model email.docx")

    context_email = {
        'nume_client': y['client']['nume'],
        'nr_cu': y['lucrare']['nr_cu'],
        'data_cu': y['lucrare']['data_cu'],
        'emitent_cu': y['lucrare']['emitent_cu'],
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nume_client': y['client']['nume'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
    }

    x.create_email(model_email, context_email, y['final_destination'])

    print("\nAvizul PMI - Spatii Verzi a fost creat \n")


def aviz_MAI(id_lucrare, path_final):
    director_final = '14.Aviz MAI'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/IS/14.Aviz MAI/'f"Model aviz MAI{' - GENERAL TEHNIC' if y['lucrare']['IDFirmaProiectare'] == 3 else ' - PROING SERV' if y['lucrare']['IDFirmaProiectare'] == 4 else ' - ROGOTEHNIC'}.docx")

    context_cerere = {
        # proiectare
        'nume_firma_proiectare': y['firma_proiectare']['nume'],
        'localitate_firma_proiectare': y['firma_proiectare']['localitate'],
        'adresa_firma_proiectare': y['firma_proiectare']['adresa'],
        'judet_firma_proiectare': y['firma_proiectare']['judet'],
        'email_firma_proiectare': y['firma_proiectare']['email'],
        'cui_firma_proiectare': y['firma_proiectare']['CUI'],
        'nr_reg_com': y['firma_proiectare']['NrRegCom'],
        'reprezentant_firma_proiectare': y['firma_proiectare']['reprezentant'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
        # client
        'nume_client': y['client']['nume'],
        'localitate_client': y['client']['localitate'],
        'adresa_client': y['client']['adresa'],
        'judet_client': y['client']['judet'],
        # beneficiar
        'nume_beneficiar': y['beneficiar']['nume'],
        # lucrare
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nr_cu': y['lucrare']['nr_cu'],
        'data_cu': y['lucrare']['data_cu'],
        'emitent_cu': y['lucrare']['emitent_cu'],
        # Data
        'data': y['astazi'],
    }

    cerere_MAI_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz MAI - DE PRINTAT.pdf")

    pdf_list = [
        cerere_MAI_pdf_path,
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CaleMemoriuTehnic'].strip('"'),
        y['lucrare']['CaleActeBeneficiar'].strip('"'),
    ]

    x.merge_pdfs_print(pdf_list, path_document_final)

    if os.path.exists(cerere_MAI_pdf_path):
        os.remove(cerere_MAI_pdf_path)

    print("\nAvizul MAI a fost creat \n")


def aviz_ISU(id_lucrare, path_final):
    director_final = '15.Aviz ISU'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/IS/15.Aviz ISU/'f"01.Cerere aviz ISU{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

    context_cerere = {
        # proiectare

        'nume_firma_proiectare': y['firma_proiectare']['nume'],
        'localitate_firma_proiectare': y['firma_proiectare']['localitate'],
        'adresa_firma_proiectare': y['firma_proiectare']['adresa'],
        'judet_firma_proiectare': y['firma_proiectare']['judet'],
        'email_firma_proiectare': y['firma_proiectare']['email'],
        'cui_firma_proiectare': y['firma_proiectare']['CUI'],
        'nr_reg_com': y['firma_proiectare']['NrRegCom'],
        'reprezentant_firma_proiectare': y['firma_proiectare']['reprezentant'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
        # client
        'nume_client': y['client']['nume'],
        'localitate_client': y['client']['localitate'],
        'adresa_client': y['client']['adresa'],
        'judet_client': y['client']['judet'],
        # beneficiar
        'nume_beneficiar': y['beneficiar']['nume'],
        # lucrare
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nr_cu': y['lucrare']['nr_cu'],
        'data_cu': y['lucrare']['data_cu'],
        'emitent_cu': y['lucrare']['emitent_cu'],
        # Data
        'data': y['astazi'],
    }

    cerere_ISU_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'])

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz ISU{y['client']['nume']} conform CU {y['lucrare']['nr_cu']} din {y['lucrare']['data_cu']}.pdf")

    # -----------------------------------------------------------------------------------------------------------------------------------------

    # creez OPISul
    model_opis = r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\IS\15.Aviz ISU\01.Opis documente.docx"
    date = x.count_pages_ISU(cerere_ISU_pdf_path, y['lucrare']['CaleCU'].strip('"'), y['lucrare']['CalePlanIncadrare'].strip(
        '"'), y['lucrare']['CalePlanSituatie'].strip('"'), y['lucrare']['CaleMemoriuTehnic'].strip('"'), y['lucrare']['CaleActeFacturare'].strip('"'))
    context_opis = {
        # lucrare
        'nr_cu': y['lucrare']['nr_cu'],
        'data_cu': y['lucrare']['data_cu'],
        # data
        'data': y['astazi'],
        # numar file
        'file_cerere': date['cerere'],
        'file_cu': date['cu'],
        'file_plan_sit': date['plan_incadrare'],
        'file_plan_inc': date['plan_situatie'],
        'file_memoriu': date['memoriu_tehnic'],
        'file_certificat': date['acte_facturare'],
    }

    opis_path = x.create_document(
        model_opis, context_opis, y['final_destination'])

    # ---------------------------------------------------------------------------------------------------------------------------------------
    pdf_list = [
        cerere_ISU_pdf_path,
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CaleMemoriuTehnic'].strip('"'),
        y['lucrare']['CaleActeFacturare'].strip('"'),
        opis_path,
    ]

    x.merge_pdfs(pdf_list, path_document_final)

    if os.path.exists(cerere_ISU_pdf_path):
        os.remove(cerere_ISU_pdf_path)
    if os.path.exists(opis_path):
        os.remove(opis_path)

    # -----------------------------------------------------------------------------------------------

    # creez EMAILUL

    model_email = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\IS\15.Aviz ISU\Model email.docx")

    context_email = {
        'nume_client': y['client']['nume'],
        'nr_cu': y['lucrare']['nr_cu'],
        'data_cu': y['lucrare']['data_cu'],
        'emitent_cu': y['lucrare']['emitent_cu'],
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nume_client': y['client']['nume'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
    }

    x.create_email(model_email, context_email, y['final_destination'])

    print("\nAvizul ISU a fost creat \n")


def aviz_Salubris(id_lucrare, path_final):
    director_final = '16.Aviz Salubris'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/IS/16.Aviz Salubris/'f"01.Cerere aviz Salubris{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

    context_cerere = {
        # proiectare
        'nume_firma_proiectare': y['firma_proiectare']['nume'],
        'localitate_firma_proiectare': y['firma_proiectare']['localitate'],
        'adresa_firma_proiectare': y['firma_proiectare']['adresa'],
        'judet_firma_proiectare': y['firma_proiectare']['judet'],
        'email_firma_proiectare': y['firma_proiectare']['email'],
        'cui_firma_proiectare': y['firma_proiectare']['CUI'],
        'nr_reg_com': y['firma_proiectare']['NrRegCom'],
        'reprezentant_firma_proiectare': y['firma_proiectare']['reprezentant'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
        # client
        'nume_client': y['client']['nume'],
        'localitate_client': y['client']['localitate'],
        'adresa_client': y['client']['adresa'],
        'judet_client': y['client']['judet'],
        # beneficiar
        'nume_beneficiar': y['beneficiar']['nume'],
        # lucrare
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nr_cu': y['lucrare']['nr_cu'],
        'data_cu': y['lucrare']['data_cu'],
        'emitent_cu': y['lucrare']['emitent_cu'],
        # Data
        'data': y['astazi'],
    }

    cerere_Salubris_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz Salubris - DE PRINTAT.pdf")

    pdf_list = [
        cerere_Salubris_pdf_path,
        y['lucrare']['CaleCU'].strip('"'),
    ]

    x.merge_pdfs_print(pdf_list, path_document_final)

    if os.path.exists(cerere_Salubris_pdf_path):
        os.remove(cerere_Salubris_pdf_path)

    print("\nAvizul Salubris a fost creat \n")


def aviz_DSP(id_lucrare, path_final):
    director_final = '17.Aviz DSP'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\IS\17.Aviz Sanatate DSP\01.Cerere aviz DSP.docx")

    context_cerere = {
        'nume_firma_proiectare': y['firma_proiectare']['nume'],
        'nume_beneficiar': y['beneficiar']['nume'],
        'nume_client': y['client']['nume'],
        'localitate_firma_proiectare': y['firma_proiectare']['localitate'],
        'adresa_firma_proiectare': y['firma_proiectare']['adresa'],
        'judet_firma_proiectare': y['firma_proiectare']['judet'],
        'email_firma_proiectare': y['firma_proiectare']['email'],
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
        'data': y['astazi'],
    }

    cerere_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    # -------------------------------------------------------------------------------------------------------------

    # creez EMAILUL

    model_email = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\IS\17.Aviz Sanatate DSP\Model email.docx")

    context_email = {
        'nume_client': y['client']['nume'],
        'nr_cu': y['lucrare']['nr_cu'],
        'data_cu': y['lucrare']['data_cu'],
        'emitent_cu': y['lucrare']['emitent_cu'],
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nume_client': y['client']['nume'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
    }

    x.create_email(model_email, context_email, y['final_destination'])

    # -----------------------------------------------------------------------------------------------

    # creez DOCUMENTUL FINAL

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz DSP Iași - {y['client']['nume']} conform CU nr. {y['lucrare']['nr_cu']} din {y['lucrare']['data_cu']}.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['lucrare']['CaleChitantaDSP'].strip('"'),
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CaleActeBeneficiar'].strip('"'),
    ]

    x.merge_pdfs(pdf_list, path_document_final)

    # -----------------------------------------------------------------------------------------------

    # fac curat

    if os.path.exists(cerere_pdf_path):
        os.remove(cerere_pdf_path)

    print("\nAvizul DSP Iași a fost creat \n")


def aviz_HCL(id_lucrare, path_final):
    director_final = '18.Aviz HCL'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/IS/18.Aviz HCL/'f"Cerere HCL{' - GENERAL TEHNIC' if y['lucrare']['IDFirmaProiectare'] == 3 else ' - PROING SERV' if y['lucrare']['IDFirmaProiectare'] == 4 else ' - ROGOTEHNIC'}.docx")

    context_cerere = {
        # proiectare
        'nume_firma_proiectare': y['firma_proiectare']['nume'],
        'localitate_firma_proiectare': y['firma_proiectare']['localitate'],
        'adresa_firma_proiectare': y['firma_proiectare']['adresa'],
        'judet_firma_proiectare': y['firma_proiectare']['judet'],
        'email_firma_proiectare': y['firma_proiectare']['email'],
        'cui_firma_proiectare': y['firma_proiectare']['CUI'],
        'nr_reg_com': y['firma_proiectare']['NrRegCom'],
        'reprezentant_firma_proiectare': y['firma_proiectare']['reprezentant'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
        # client
        'nume_client': y['client']['nume'],
        'localitate_client': y['client']['localitate'],
        'adresa_client': y['client']['adresa'],
        'judet_client': y['client']['judet'],
        # beneficiar
        'nume_beneficiar': y['beneficiar']['nume'],
        # lucrare
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nr_cu': y['lucrare']['nr_cu'],
        'data_cu': y['lucrare']['data_cu'],
        'emitent_cu': y['lucrare']['emitent_cu'],
        'suprafata_mp': y['lucrare']['SuprafataMP'],
        'lungime_metri': y['lucrare']['LungimeTraseuMetri'],
        # Data
        'data': y['astazi'],
    }

    cerere_HCL_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz HCL - DE PRINTAT.pdf")

    pdf_list = [
        cerere_HCL_pdf_path,
        y['lucrare']['CaleAvizGiS'].strip('"'),
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CaleMemoriuTehnic'].strip('"'),
        y['lucrare']['CaleAvizCTEsauATR'].strip('"'),
        y['lucrare']['CaleActeBeneficiar'].strip('"'),
    ]


    with os.scandir(y['lucrare']['CaleExtraseCF'].strip('"')) as entries:
        for entry in entries:
            if entry.is_file() and "Extras" in str(entry):
                pdf_list.append(entry.path)

    x.merge_pdfs_print(pdf_list, path_document_final)

    if os.path.exists(cerere_HCL_pdf_path):
        os.remove(cerere_HCL_pdf_path)

    print("\nAvizul HCL a fost creat \n")


def aviz_PMI_GiS(id_lucrare, path_final):
    director_final = '19.Aviz PMI - GiS'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/IS/19.Aviz PMI - GiS/'f"Cerere Aviz GiS{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

    context_cerere = {
        # proiectare

        'nume_firma_proiectare': y['firma_proiectare']['nume'],
        'localitate_firma_proiectare': y['firma_proiectare']['localitate'],
        'adresa_firma_proiectare': y['firma_proiectare']['adresa'],
        'judet_firma_proiectare': y['firma_proiectare']['judet'],
        'email_firma_proiectare': y['firma_proiectare']['email'],
        'cui_firma_proiectare': y['firma_proiectare']['CUI'],
        'nr_reg_com': y['firma_proiectare']['NrRegCom'],
        'reprezentant_firma_proiectare': y['firma_proiectare']['reprezentant'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
        # client
        'nume_client': y['client']['nume'],
        'localitate_client': y['client']['localitate'],
        'adresa_client': y['client']['adresa'],
        'judet_client': y['client']['judet'],
        # beneficiar
        'nume_beneficiar': y['beneficiar']['nume'],
        # lucrare
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nr_cu': y['lucrare']['nr_cu'],
        'data_cu': y['lucrare']['data_cu'],
        'emitent_cu': y['lucrare']['emitent_cu'],
        # Data
        'data': y['astazi'],
    }

    cerere_PMI_GiS_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz Cadastru - GiS - {y['client']['nume']} conform CU {y['lucrare']['nr_cu']} din {y['lucrare']['data_cu']}.pdf")

    pdf_list = [
        cerere_PMI_GiS_pdf_path,
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CaleACCUConstructie'].strip('"'),
        y['lucrare']['CaleActeBeneficiar'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CaleMemoriuTehnic'].strip('"'),

    ]


    with os.scandir(y['lucrare']['CaleExtraseCF'].strip('"')) as entries:
        for entry in entries:
            if entry.is_file() and "Extras" in str(entry):
                pdf_list.append(entry.path)


    x.merge_pdfs(pdf_list, path_document_final)

    if os.path.exists(cerere_PMI_GiS_pdf_path):
        os.remove(cerere_PMI_GiS_pdf_path)

    x.copy_file(y['lucrare']['CaleRidicareTopoDWG'], path_final,
                director_final, 'Ridicare TOPO.dwg')
    # -----------------------------------------------------------------------------------------------

    # creez EMAILUL

    model_email = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\IS\19.Aviz PMI - GiS\Model email.docx")

    context_email = {
        'nume_client': y['client']['nume'],
        'nr_cu': y['lucrare']['nr_cu'],
        'data_cu': y['lucrare']['data_cu'],
        'emitent_cu': y['lucrare']['emitent_cu'],
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nume_client': y['client']['nume'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
    }

    x.create_email(model_email, context_email, y['final_destination'])

    print("\nAvizul PMI - GiS a fost creat \n")


def aviz_Nomenclatura(id_lucrare, path_final):
    director_final = '20.Certificat nomenclatura urbana'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\IS\20.Nomenclatura urbana\Cerere nomenclatura urbana.docx")

    context_cerere = {
        'nume_firma_proiectare': y['firma_proiectare']['nume'],
        'nume_beneficiar': y['beneficiar']['nume'],
        'nume_client': y['client']['nume'],
        'localitate_firma_proiectare': y['firma_proiectare']['localitate'],
        'adresa_firma_proiectare': y['firma_proiectare']['adresa'],
        'judet_firma_proiectare': y['firma_proiectare']['judet'],
        'email_firma_proiectare': y['firma_proiectare']['email'],
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
        'data': y['astazi'],
    }

    cerere_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    # creez EMAILUL

    model_email = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\IS\20.Nomenclatura urbana\Model email.docx")

    context_email = {
        'nume_client': y['client']['nume'],
        'nr_cu': y['lucrare']['nr_cu'],
        'data_cu': y['lucrare']['data_cu'],
        'emitent_cu': y['lucrare']['emitent_cu'],
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nume_client': y['client']['nume'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
    }

    x.create_email(model_email, context_email, y['final_destination'])

    # -----------------------------------------------------------------------------------------------

    # creez DOCUMENTUL FINAL

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie Cerere nomenclatura urbana - {y['client']['nume']} conform CU {y['lucrare']['nr_cu']} din {y['lucrare']['data_cu']}.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CaleActeBeneficiar'].strip('"'),
    ]


    with os.scandir(y['lucrare']['CaleExtraseCF'].strip('"')) as entries:
        for entry in entries:
            if entry.is_file() and "Extras" in str(entry):
                pdf_list.append(entry.path)

    x.merge_pdfs_print(pdf_list, path_document_final)

    # -----------------------------------------------------------------------------------------------

    # fac curat

    if os.path.exists(cerere_pdf_path):
        os.remove(cerere_pdf_path)

    print("\nAvizul Nomenclatura urbana a fost creat \n")
