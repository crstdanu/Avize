import os
import functii as x


def aviz_APM(id_lucrare, path_final):
    director_final = '01.Mediu APM Neamt'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere_APM = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/NT/01.Mediu APM - Neamt/'f"01.Cerere{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

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
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\NT\01.Mediu APM - Neamt\02.Notificare.docx")

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
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\NT\01.Mediu APM - Neamt\Model email.docx")

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
        path_final, director_final, f"Documentatie aviz APM Neamț - {y['client']['nume']} conform CU {y['lucrare']['nr_cu']} din {y['lucrare']['data_cu']}.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['lucrare']['CaleChitantaAPM'].strip('"'),
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        notificare_pdf_path,
        y['lucrare']['CaleActeFacturare'].strip('"'),
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

    print("\nAvizul APM Neamț a fost creat \n")


def aviz_EE_Delgaz(id_lucrare, path_final):
    director_final = '02.Aviz EE Delgaz - Neamt'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere_EE_Delgaz = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/NT/02.Aviz EE Delgaz - Neamt/'f"01.Cerere aviz EE Delgaz{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

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

    # PUNEM PLANUL DE SITUATIE DETALIAT
    x.copy_file(y['lucrare']['CalePlanSituatiePDF'], path_final,
                director_final, 'Plan situatie.pdf')
    x.copy_file(y['lucrare']['CalePlanSituatieDWG'], path_final,
                director_final, 'Plan situatie.dwg')

    # -----------------------------------------------------------------------------------------------

    # creez EMAILUL

    model_email = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/NT/02.Aviz EE Delgaz - Neamt/'f"Model email{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

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

    print("\nAvizul EE Delgaz - Neamț a fost creat \n")


def aviz_GN_Delgaz(id_lucrare, path_final):
    director_final = '03.Aviz GN Delgaz - Neamt'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/NT/03.Aviz GN Delgaz/'f"01.Aviz GN{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

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

    cerere_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz GN - Delgaz - {y['client']['nume']} conform CU nr. {y['lucrare']['nr_cu']} din {y['lucrare']['data_cu']}.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CaleMemoriuTehnic'].strip('"'),
        y['lucrare']['CaleActeFacturare'].strip('"'),
    ]

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
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/NT/03.Aviz GN Delgaz/'f"Model email{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

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

    print("\nAvizul GN Delgaz - Neamț a fost creat \n")


def aviz_ApaServ(id_lucrare, path_final):
    director_final = '04.Aviz ApaServ - Neamt'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\NT\04.Aviz ApaServ\Cerere aviz ApaServ.docx")

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
        # REPREZENTANT
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
        path_final, director_final, f"Documentatie aviz ApaServ - {y['client']['nume']} conform CU nr. {y['lucrare']['nr_cu']} din {y['lucrare']['data_cu']}.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CaleMemoriuTehnic'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CaleActeFacturare'].strip('"'),
        y['firma_proiectare']['caleCI'].strip('"'),
    ]

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
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\NT\04.Aviz ApaServ\Model email.docx")

    context_email = {
        'cui_firma_proiectare': y['firma_proiectare']['CUI'],
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

    print("\nAvizul ApaServ - Neamț a fost creat \n")


def aviz_SGA(id_lucrare, path_final):
    director_final = '05.Aviz SGA - Neamt'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\NT\05.Gospodarirea Apelor Neamt\01.Model aviz SGA-Neamt.docx")

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

    cerere_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz SGA - Neamt - {y['client']['nume']} conform CU nr. {y['lucrare']['nr_cu']} din {y['lucrare']['data_cu']}.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CaleMemoriuTehnic'].strip('"'),
        y['lucrare']['CaleActeFacturare'].strip('"'),
    ]

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
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\NT\05.Gospodarirea Apelor Neamt\Model email.docx")

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

    print("\nAvizul SGA - Neamț a fost creat \n")


def aviz_Luxten(id_lucrare, path_final):
    director_final = '06.Aviz Luxten - Neamt'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/NT/06.Aviz Luxten/'f"Cerere Aviz Luxten{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

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

    cerere_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz Luxten - {y['client']['nume']} conform CU nr. {y['lucrare']['nr_cu']} din {y['lucrare']['data_cu']}.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CaleMemoriuTehnic'].strip('"'),
        y['lucrare']['CaleActeFacturare'].strip('"'),
    ]

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
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\NT\06.Aviz Luxten\Model email.docx")
    
    facturare = x.facturare(id_lucrare)

    context_email = {
        'cui_firma_proiectare': y['firma_proiectare']['CUI'],
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
        'firma_facturare': facturare['firma_facturare'],
        'cui_firma_facturare': facturare['cui_firma_facturare'],
    }

    x.create_email(model_email, context_email, y['final_destination'])

    print("\nAvizul Luxten - Neamț a fost creat \n")


def aviz_Orange(id_lucrare, path_final):
    director_final = '07.Aviz Orange - Neamt'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/NT/07.Aviz Orange/'f"Cerere Orange{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

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

    cerere_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))
    os.rename(cerere_pdf_path, os.path.join(
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
    x.copy_file("G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/NT/07.Aviz Orange/Citeste-ma.docx",
                path_final, director_final, 'Citeste-ma.docx')
    # -----------------------------------------------------------------------------------------------

    print("\nAvizul Orange a fost creat \n")


def aviz_PMPN_Trafic(id_lucrare, path_final):
    director_final = '08.Aviz PMPN - Trafic'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\NT\08.Aviz PMPN - Trafic\01.Model aviz PMPN Trafic.docx")

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

    cerere_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz PMPN - Trafic - {y['client']['nume']} conform CU nr. {y['lucrare']['nr_cu']} din {y['lucrare']['data_cu']}.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CaleMemoriuTehnic'].strip('"'),
        y['lucrare']['CaleActeFacturare'].strip('"'),
    ]

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
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\NT\08.Aviz PMPN - Trafic\Model email.docx")

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

    print("\nAvizul PMPN - Trafic a fost creat \n")


def aviz_PMPN_Protocol_HCL(id_lucrare, path_final):
    director_final = '09.Aviz PMPN - Protocol HCL'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\NT\09. Aviz PMPN - Protocol HCL\01.Model aviz PMPN Protocol HCL.docx")

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

    cerere_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz PMPN - Protocol - {y['client']['nume']} conform CU nr. {y['lucrare']['nr_cu']} din {y['lucrare']['data_cu']}.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CaleMemoriuTehnic'].strip('"'),
        y['lucrare']['CaleActeFacturare'].strip('"'),
    ]

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
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\NT\09. Aviz PMPN - Protocol HCL\Model email.docx")

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

    print("\nAvizul PMPN - Protocol HCL a fost creat \n")


def aviz_Publiserv(id_lucrare, path_final):
    director_final = '10.Aviz Publiserv - Neamt'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\NT\10.Aviz Publiserv\01.Model aviz Publiserv.docx")

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

    cerere_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz Publiserv - {y['client']['nume']} conform CU nr. {y['lucrare']['nr_cu']} din {y['lucrare']['data_cu']}.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CaleMemoriuTehnic'].strip('"'),
        y['lucrare']['CaleActeFacturare'].strip('"'),
    ]

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
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\NT\10.Aviz Publiserv\Model email.docx")

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

    print("\nAvizul Publiserv - Neamt a fost creat \n")


def aviz_TransGaz(id_lucrare, path_final):
    director_final = '11.Aviz TransGaz - Neamt'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/NT/11. Aviz TransGaz/'f"Cerere aviz TransGaz{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

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

    cerere_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz TransGaz - {y['client']['nume']} conform CU {y['lucrare']['nr_cu']} din {y['lucrare']['data_cu']}.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CaleMemoriuTehnic'].strip('"'),
        y['lucrare']['CaleActeFacturare'].strip('"'),
    ]

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
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\NT\11. Aviz TransGaz\Model email.docx")
    
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
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
        'firma_facturare': facturare['firma_facturare'],
        'cui_firma_facturare': facturare['cui_firma_facturare'],
        

    }

    x.create_email(model_email, context_email, y['final_destination'])

    print("\nAvizul TransGaz - Neamț a fost creat \n")


def aviz_Cultura(id_lucrare, path_final):
    director_final = '12.Aviz Cultura - Neamt'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/NT/12. Aviz Cultura - Neamt/'f"Cerere Cultura Neamt{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

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

    cerere_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz Culrura - {y['client']['nume']} conform CU {y['lucrare']['nr_cu']} din {y['lucrare']['data_cu']}.pdf")

    path_document_printabil = os.path.join(
        path_final, director_final, f"Documentatie aviz Culrura - DE PRINTAT.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CaleMemoriuTehnic'].strip('"'),
        y['lucrare']['CaleActeFacturare'].strip('"'),
    ]

    x.merge_pdfs(pdf_list, path_document_final)
    x.merge_pdfs_print(pdf_list, path_document_printabil)

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
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\NT\12. Aviz Cultura - Neamt\Model email.docx")

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

    print("\nAvizul Cultura - Neamț a fost creat \n")


def aviz_CFR(id_lucrare, path_final):
    director_final = '18.Aviz CFR'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/NT/18. Aviz CFR/'f"Cerere aviz CFR{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

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

    cerere_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz CFR - DE PRINTAT.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['lucrare']['CaleCU'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanIncadrare'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CalePlanSituatie'].strip('"'),
        y['lucrare']['CaleMemoriuTehnic'].strip('"'),
        y['lucrare']['CaleActeFacturare'].strip('"'),
    ]

    x.merge_pdfs_print(pdf_list, path_document_final)

    if os.path.exists(cerere_pdf_path):
        os.remove(cerere_pdf_path)

    print("\nAvizul CFR a fost creat \n")


def aviz_HCL(id_lucrare, path_final):
    director_final = '19.Aviz HCL'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/NT/19.Aviz HCL/'f"Cerere HCL{' - GENERAL TEHNIC' if y['lucrare']['IDFirmaProiectare'] == 3 else ' - PROING SERV' if y['lucrare']['IDFirmaProiectare'] == 4 else ' - ROGOTEHNIC'}.docx")

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

    with os.scandir(y['lucrare']['CaleExtraseCF'].strip('"')) as entries:
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