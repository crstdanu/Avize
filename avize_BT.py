import os
import functii as x



def aviz_APM(id_lucrare, path_final):
    director_final = '01.Mediu APM Botosani'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere_APM = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/BT/01.Mediu APM/'f"01.Cerere{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

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
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\BT\01.Mediu APM\02.Notificare.docx")

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
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\BT\01.Mediu APM\Model email.docx")

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
        path_final, director_final, f"Documentatie aviz APM Botoșani - {y['client']['nume']} conform CU {y['lucrare']['nr_cu']} din {y['lucrare']['data_cu']}.pdf")

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


    # PUNEM PLANUL DE SITUATIE DETALIAT
    x.copy_file(y['lucrare']['CalePlanSituatiePDF'], path_final,
                director_final, 'Plan situatie.pdf')
    x.copy_file(y['lucrare']['CalePlanSituatieDWG'], path_final,
                director_final, 'Plan situatie.dwg')

    # -----------------------------------------------------------------------------------------------

    # fac curat

    if os.path.exists(cerere_pdf_path):
        os.remove(cerere_pdf_path)

    if os.path.exists(notificare_pdf_path):
        os.remove(notificare_pdf_path)

    print("\nAvizul APM Botoșani a fost creat \n")


def aviz_EE_Delgaz(id_lucrare, path_final):
    director_final = '02.Aviz EE Delgaz - Botoșani'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere_EE_Delgaz = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/BT/02.Aviz EE Delgaz/'f"01.Cerere aviz EE Delgaz{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

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

    # PUNEM PLANUL DE SITUATIE DETALIAT
    x.copy_file(y['lucrare']['CalePlanSituatiePDF'], path_final,
                director_final, 'Plan situatie.pdf')
    x.copy_file(y['lucrare']['CalePlanSituatieDWG'], path_final,
                director_final, 'Plan situatie.dwg')

    if os.path.exists(cerere_EE_pdf_path):
        os.remove(cerere_EE_pdf_path)

    # -----------------------------------------------------------------------------------------------

    # creez EMAILUL

    model_email = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/BT/02.Aviz EE Delgaz/'f"Model email{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

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

    print("\nAvizul EE Delgaz - Botosani a fost creat \n")


def aviz_GN_Delgaz(id_lucrare, path_final):
    director_final = '03.Aviz GN Delgaz - Botoșani'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/BT/03.Aviz GN Delgaz/'f"01.Aviz GN{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

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


        # PUNEM PLANUL DE SITUATIE DETALIAT
    x.copy_file(y['lucrare']['CalePlanSituatiePDF'], path_final,
                director_final, 'Plan situatie.pdf')
    x.copy_file(y['lucrare']['CalePlanSituatieDWG'], path_final,
                director_final, 'Plan situatie.dwg')

    if os.path.exists(cerere_GN_pdf_path):
        os.remove(cerere_GN_pdf_path)

    # -----------------------------------------------------------------------------------------------

    # creez EMAILUL

    model_email = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/BT/03.Aviz GN Delgaz/'f"Model email{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

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

    print("\nAvizul GN Delgaz - Bacău a fost creat \n")


def aviz_Orange(id_lucrare, path_final):
    director_final = '07.Aviz Orange'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/BT/07.Aviz Orange/'f"Cerere Orange{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

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
    x.copy_file("G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/BT/07.Aviz Orange/Citeste-ma.docx",
                path_final, director_final, 'Citeste-ma.docx')
    # -----------------------------------------------------------------------------------------------

    print("\nAvizul Orange a fost creat \n")


def aviz_HCL(id_lucrare, path_final):
    director_final = '18.Aviz HCL'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/BT/18.Aviz HCL/'f"Cerere HCL{' - GENERAL TEHNIC' if y['lucrare']['IDFirmaProiectare'] == 3 else ' - PROING SERV' if y['lucrare']['IDFirmaProiectare'] == 4 else ' - ROGOTEHNIC'}.docx")

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
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CaleMemoriuTehnic'].strip('"'),
        y['lucrare']['CaleActeFacturare'].strip('"'),
        ]

    with os.scandir(y['lucrare']['CaleExtraseCF']) as entries:
        for entry in entries:
            if entry.is_file() and "Extras" in str(entry):
                pdf_list.append(entry.path)

    x.merge_pdfs(pdf_list, path_document_final)

    # PUNEM PLANUL DE SITUATIE DETALIAT
    x.copy_file(y['lucrare']['CalePlanSituatiePDF'], path_final,
                director_final, 'Plan situatie.pdf')
    x.copy_file(y['lucrare']['CalePlanSituatieDWG'], path_final,
                director_final, 'Plan situatie.dwg')

    if os.path.exists(cerere_HCL_pdf_path):
        os.remove(cerere_HCL_pdf_path)

    print("\nAvizul HCL a fost creat \n")


def aviz_SGA(id_lucrare, path_final):
    director_final = '05.Aviz SGA - Botosani'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/BT/05. Aviz SGA/Cerere aviz SGA'f"{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

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
        # reprezentant
        'localitate_repr': y['firma_proiectare']['localitate_repr'],
        'adresa_repr': y['firma_proiectare']['adresa_repr'],
        'judet_repr': y['firma_proiectare']['judet_repr'],
        'seria_CI': y['firma_proiectare']['seria_CI'],
        'nr_CI': y['firma_proiectare']['nr_CI'],
        'data_CI': y['firma_proiectare']['data_CI'],
        'cnp_repr': y['firma_proiectare']['cnp_repr'],
        # Data
        'data': y['astazi'],
    }

    cerere_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz SGA - Botoșani - {y['client']['nume']} conform CU nr. {y['lucrare']['nr_cu']} din {y['lucrare']['data_cu']}.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CaleMemoriuTehnic'].strip('"'),
        y['lucrare']['CaleActeFacturare'].strip('"'),
    ]

    with os.scandir(y['lucrare']['CaleExtraseCF'].strip('"')) as entries:
        for entry in entries:
            if entry.is_file() and "Extras" in str(entry):
                pdf_list.append(entry.path)

    x.merge_pdfs(pdf_list, path_document_final)

    if os.path.exists(cerere_pdf_path):
        os.remove(cerere_pdf_path)

    # PUNEM PLANUL DE SITUATIE DETALIAT
    x.copy_file(y['lucrare']['CalePlanSituatiePDF'], path_final,
                director_final, 'Plan situatie.pdf')
    x.copy_file(y['lucrare']['CalePlanSituatieDWG'], path_final,
                director_final, 'Plan situatie.dwg')

    # -----------------------------------------------------------------------------------------------

    # creez EMAILUL

    model_email = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\BT\05. Aviz SGA\Model email.docx")

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

    print("\nAvizul SGA - Botosani a fost creat \n")