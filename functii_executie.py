import os
import functii as x


def conventie_lucrari(id_executie, path_final):
    director_final = '01.Conventie lucrari'
    cale_stampila = "G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/DOCUMENTE/Stampila - RGT.png"
    y = x.get_data_executie(path_final, director_final, id_executie)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\Executie\01. Pentru incepere lucrari\00. Cerere CL.docx")

    context_cerere = {
        'nr_cl': y['executie']['nr_cl'],
        'nume_lucrare': y['executie']['nume'],
        'localitate_lucrare': y['executie']['localitate'],
        'adresa_lucrare': y['executie']['adresa'],
        'judet_lucrare': y['executie']['judet'],
        'nume_client': y['client']['nume'],
        # Data
        'data': y['astazi'],
    }

    cerere_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], cale_stampila)

    x.move_file(cerere_pdf_path, path_final,
                director_final, f"00.Cerere CL.pdf")

    x.copy_file(y['executie']['CaleInstruireColectiva'], path_final,
                director_final, '02.SCAN - Instruire colectiva.pdf')

    x.copy_file_prefix(y['executie']['CaleContract1'],
                       path_final, director_final, '03.')

    x.copy_file_prefix(y['executie']['CaleContract2'],
                       path_final, director_final, '03.')

    if y['executie']['CaleContract3']:
        x.copy_file_prefix(y['executie']['CaleContract3'],
                           path_final, director_final, '03.')

    x.copy_file_prefix(y['executie']['CaleAvizCTEsauATR'],
                       path_final, director_final, '04.')

    x.copy_file(y['executie']['CaleMemoriuTehnicACScanat'], path_final,
                director_final, '05. Memoriu tehnic PTH.pdf')

    x.copy_file(y['executie']['CalePlanIncadrarePTH'], path_final,
                director_final, '06. Plan incadrare PTH.pdf')

    x.copy_file(y['executie']['CalePlanSituatiePTH'], path_final,
                director_final, '07. Plan situatie PTH.pdf')

    x.copy_file(y['executie']['CaleSchemaMonofilaraJT'], path_final,
                director_final, '08. Schema monofilara JT.pdf')

    if y['executie']['CaleSchemaMonofilaraMT']:
        x.copy_file(y['executie']['CaleSchemaMonofilaraMT'],
                    path_final, director_final, '09. Schema monofilara MT.pdf')

    path_document_final = os.path.join(
        path_final, director_final, f"10. AC+planse.pdf")

    pdf_list = [
        y['executie']['CaleACScanat'],
        y['executie']['CalePlanIncadrareACScanat'],
        y['executie']['CalePlanSituatieACScanat'],

    ]

    x.merge_pdfs_print(pdf_list, path_document_final)

    # -----------------------------------------------------------------------------------------------

    # creez EMAILUL

    model_email = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/IS/02.Aviz EE Delgaz - Iasi/'f"Model email{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

    facturare = x.facturare(id_executie)

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
        'firma_facturare': facturare['firma_facturare'],
        'cui_firma_facturare': facturare['cui_firma_facturare'],
    }

    x.create_email(model_email, context_email, y['final_destination'])

    print("\nConvenția de lucrări a fost creată \n")
