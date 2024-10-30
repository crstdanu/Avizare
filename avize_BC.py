import os
import functii as x



def aviz_APM(id_lucrare, path_final):
    director_final = '01.Mediu APM Bacau'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere_APM = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/BC/01.Mediu APM/'f"01.Cerere{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

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

    model_Notificare_APM = (r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\BC\01.Mediu APM\02.Notificare.docx")

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
        'descrierea_proiectului': y['tblCU']['descrierea_proiectului'],
        'intocmit': y['intocmit'],
        'verificat': y['verificat'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
    }

    notificare_pdf_path = x.create_document(
        model_Notificare_APM, context_Notificare_APM, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    # -----------------------------------------------------------------------------------------------

    # creez EMAILUL

    model_Email = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\BC\01.Mediu APM\Model email.docx")

    context_Email = {
        'nume_client': y['client']['nume'],
        'nr_cu': y['tblCU']['NrCU'],
        'data_cu': x.get_date(y['tblCU']['DataCU']),
        'emitent_cu': y['EmitentCU']['denumire_institutie'],

        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],

        'nume_client': y['client']['nume'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
    }

    x.create_email(model_Email, context_Email, y['final_destination'])

    # -----------------------------------------------------------------------------------------------

    # creez DOCUMENTUL FINAL

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz APM Bacău - DE PRINTAT.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['tblCU']['CaleChitantaAPM'].strip('"'),
        y['tblCU']['CaleCU'].strip('"'),
        y['tblCU']['CalePlanIncadrareCU'].strip('"'),
        y['tblCU']['CalePlanSituatieCU'].strip('"'),
        notificare_pdf_path,
    ]

    x.merge_pdfs_print(pdf_list, path_document_final)

    # PUNEM PLANUL DE SITUATIE DETALIAT
    x.copy_file(y['tblCU']['CalePlanSituatiePDF'], path_final, director_final, 'Plan situatie.pdf')
    x.copy_file(y['tblCU']['CalePlanSituatieDWG'], path_final, director_final, 'Plan situatie.dwg')

    # -----------------------------------------------------------------------------------------------

    # fac curat

    if os.path.exists(cerere_pdf_path):
        os.remove(cerere_pdf_path)

    if os.path.exists(notificare_pdf_path):
        os.remove(notificare_pdf_path)

    print("\nAvizul APM Bacău a fost creat \n")



def aviz_EE_Delgaz(id_lucrare, path_final):
    director_final = '02.Aviz EE Delgaz - Bacau'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/BC/02.Aviz EE Delgaz/'f"01.Cerere aviz EE Delgaz{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

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

        'nr_cu': y['tblCU']['NrCU'],
        'data_cu': x.get_date(y['tblCU']['DataCU']),
        'emitent_cu': y['EmitentCU']['denumire_institutie'],
        # Data
        'data': y['astazi'],
    }

    cerere_EE_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz EE - Delgaz - {y['client']['nume']} conform CU {y['tblCU']['NrCU']} din {y['lucrare']['data_cu']}.pdf")


    pdf_list = [
        cerere_EE_pdf_path,
        y['tblCU']['CaleCU'].strip('"'),
        y['tblCU']['CalePlanIncadrareCU'].strip('"'),
        y['tblCU']['CalePlanSituatieCU'].strip('"'),
        y['tblCU']['CaleMemoriuTehnicSS'].strip('"'),
        y['tblCU']['CaleActeFacturare'].strip('"'),
    ]

    x.merge_pdfs(pdf_list, path_document_final)

    # PUNEM PLANUL DE SITUATIE DETALIAT
    x.copy_file(y['tblCU']['CalePlanSituatiePDF'], path_final, director_final, 'Plan situatie.pdf')
    x.copy_file(y['tblCU']['CalePlanSituatieDWG'], path_final, director_final, 'Plan situatie.dwg')

    if os.path.exists(cerere_EE_pdf_path):
        os.remove(cerere_EE_pdf_path)

    # -----------------------------------------------------------------------------------------------

    # creez EMAILUL

    model_email = ('G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/BC/02.Aviz EE Delgaz/'f"Model email{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

    facturare = x.facturare(id_lucrare)

    context_email = {
        'nume_client': y['client']['nume'],
        'nr_cu': y['tblCU']['NrCU'],
        'data_cu': x.get_date(y['tblCU']['DataCU']),
        'emitent_cu': y['EmitentCU']['denumire_institutie'],
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

    print("\nAvizul EE Delgaz - Bacau a fost creat \n")



def aviz_GN_Delgaz(id_lucrare, path_final):
    director_final = '03.Aviz GN Delgaz - Bacau'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/BC/03.Aviz GN Delgaz/'f"01.Aviz GN{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

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

        'nr_cu': y['tblCU']['NrCU'],
        'data_cu': x.get_date(y['tblCU']['DataCU']),
        'emitent_cu': y['EmitentCU']['denumire_institutie'],
        # Data
        'data': y['astazi'],
    }

    cerere_GN_pdf_path = x.create_document(model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz GN - Delgaz - {y['client']['nume']} conform CU nr. {y['tblCU']['NrCU']} din {y['lucrare']['data_cu']}.pdf")

    pdf_list = [
        cerere_GN_pdf_path,
        y['tblCU']['CaleCU'].strip('"'),
        y['tblCU']['CalePlanIncadrareCU'].strip('"'),
        y['tblCU']['CalePlanSituatieCU'].strip('"'),
        y['tblCU']['CaleMemoriuTehnicSS'].strip('"'),
        y['tblCU']['CaleActeFacturare'].strip('"'),
    ]

    x.merge_pdfs(pdf_list, path_document_final)

    # PUNEM PLANUL DE SITUATIE DETALIAT
    x.copy_file(y['tblCU']['CalePlanSituatiePDF'], path_final, director_final, 'Plan situatie.pdf')
    x.copy_file(y['tblCU']['CalePlanSituatieDWG'], path_final, director_final, 'Plan situatie.dwg')

    if os.path.exists(cerere_GN_pdf_path):
        os.remove(cerere_GN_pdf_path)

    # -----------------------------------------------------------------------------------------------

    # creez EMAILUL

    model_email = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/BC/03.Aviz GN Delgaz/'f"Model email{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

    facturare = x.facturare(id_lucrare)

    context_email = {
        'nume_client': y['client']['nume'],
        'nr_cu': y['tblCU']['nr_cu'],
        'data_cu': y['tblCU']['data_cu'],
        'emitent_cu': y['EmitentCU']['denumire_institutie'],
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
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/BC/07.Aviz Orange/'f"Cerere Orange{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

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
        'nr_cu': y['tblCU']['NrCU'],
        'data_cu': x.get_date(y['tblCU']['DataCU']),
        'emitent_cu': y['EmitentCU']['denumire_institutie'],
        # Data
        'data': y['astazi'],
    }

    cerere_Orange_pdf_path = x.create_document(model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))
    os.rename(cerere_Orange_pdf_path, os.path.join(path_final, director_final, '01.Cerere.pdf'))

    x.copy_file(y['tblCU']['CaleCU'], path_final, director_final, '02.Certificat de urbanism.pdf')
    x.copy_file(y['tblCU']['CalePlanIncadrareCU'], path_final, director_final, '03.Plan incadrare in zona.pdf')
    x.copy_file(y['tblCU']['CalePlanSituatieCU'], path_final, director_final, '04.Plan situatie.pdf')
    x.copy_file(y['tblCU']['CaleMemoriuTehnicSS'], path_final, director_final, '05.Memoriu tehnic.pdf')
    x.copy_file(y['tblCU']['CaleActeFacturare'], path_final, director_final, '06.Acte facturare.pdf')
    x.copy_file("G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/BC/07.Aviz Orange/Citeste-ma.docx", path_final, director_final, 'Citeste-ma.docx')
    # -----------------------------------------------------------------------------------------------

    print("\nAvizul Orange a fost creat \n")



def aviz_HCL(id_lucrare, path_final):
    director_final = '18.Aviz HCL'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/BC/18.Aviz HCL/'f"Cerere HCL{' - GENERAL TEHNIC' if y['lucrare']['IDFirmaProiectare'] == 3 else ' - PROING SERV' if y['lucrare']['IDFirmaProiectare'] == 4 else ' - ROGOTEHNIC'}.docx")

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
        'nr_cu': y['tblCU']['NrCU'],
        'data_cu': x.get_date(y['tblCU']['DataCU']),
        'emitent_cu': y['EmitentCU']['denumire_institutie'],
        'suprafata_mp': y['tblCU']['SuprafataOcupata'],
        'lungime_metri': y['tblCU']['LungimeTraseu'],
        # Data
        'data': y['astazi'],
    }

    cerere_HCL_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz HCL  - {y['client']['nume']} conform CU nr. {y['tblCU']['NrCU']} din {y['lucrare']['data_cu']}.pdf.pdf")

    pdf_list = [
        cerere_HCL_pdf_path,
        y['tblCU']['CaleCU'].strip('"'),
        y['tblCU']['CalePlanIncadrareCU'].strip('"'),
        y['tblCU']['CalePlanSituatieCU'].strip('"'),
        y['tblCU']['CaleMemoriuTehnicSS'].strip('"'),
        y['tblCU']['CaleActeFacturare'].strip('"'),
    ]

    with os.scandir(y['tblCU']['CaleExtraseCF'].strip('"')) as entries:
        for entry in entries:
            if entry.is_file() and "Extras" in str(entry):
                pdf_list.append(entry.path)

    x.merge_pdfs(pdf_list, path_document_final)

    # PUNEM PLANUL DE SITUATIE DETALIAT
    x.copy_file(y['tblCU']['CalePlanSituatiePDF'], path_final, director_final, 'Plan situatie.pdf')
    x.copy_file(y['tblCU']['CalePlanSituatieDWG'], path_final, director_final, 'Plan situatie.dwg')

    if os.path.exists(cerere_HCL_pdf_path):
        os.remove(cerere_HCL_pdf_path)

    print("\nAvizul HCL a fost creat \n")



def aviz_SGA(id_lucrare, path_final):
    director_final = '05.Aviz SGA - Bacau'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/BC/05. Aviz SGA/Cerere aviz SGA'f"{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

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
        'nr_cu': y['tblCU']['NrCU'],
        'data_cu': x.get_date(y['tblCU']['DataCU']),
        'emitent_cu': y['EmitentCU']['denumire_institutie'],
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
        path_final, director_final, f"Documentatie aviz SGA - Neamt - {y['client']['nume']} conform CU nr. {y['tblCU']['NrCU']} din {x.get_date(y['tblCU']['DataCU'])}.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['tblCU']['CaleCU'].strip('"'),
        y['tblCU']['CalePlanIncadrareCU'].strip('"'),
        y['tblCU']['CalePlanSituatieCU'].strip('"'),
        y['tblCU']['CaleMemoriuTehnicSS'].strip('"'),
        y['tblCU']['CaleActeFacturare'].strip('"'),
    ]

    x.merge_pdfs(pdf_list, path_document_final)

    if os.path.exists(cerere_pdf_path):
        os.remove(cerere_pdf_path)

    # PUNEM PLANUL DE SITUATIE DETALIAT
    x.copy_file(y['tblCU']['CalePlanSituatiePDF'], path_final, director_final, 'Plan situatie.pdf')
    x.copy_file(y['tblCU']['CalePlanSituatieDWG'], path_final, director_final, 'Plan situatie.dwg')

    # -----------------------------------------------------------------------------------------------

    # creez EMAILUL

    model_email = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\BC\05. Aviz SGA\Model email.docx")

    context_email = {
        'nume_client': y['client']['nume'],
        'nr_cu': y['tblCU']['NrCU'],
        'data_cu': x.get_date(y['tblCU']['DataCU']),
        'emitent_cu': y['EmitentCU']['denumire_institutie'],
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nume_client': y['client']['nume'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
    }

    x.create_email(model_email, context_email, y['final_destination'])

    print("\nAvizul SGA - Bacau a fost creat \n")


def aviz_MApN(id_lucrare, path_final):
    director_final = '09.Aviz MApN'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/BC/09.Aviz MApN/'f"Cerere MApN{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

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
        'nr_cu': y['tblCU']['NrCU'],
        'data_cu': x.get_date(y['tblCU']['DataCU']),
        'emitent_cu': y['EmitentCU']['denumire_institutie'],
        # Data
        'data': y['astazi'],
    }

    cerere_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz MApN - DE PRINTAT.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['tblCU']['CaleCU'].strip('"'),
        y['tblCU']['CalePlanIncadrareCU'].strip('"'),
        y['tblCU']['CalePlanIncadrareCU'].strip('"'),
        y['tblCU']['CalePlanSituatieCU'].strip('"'),
        y['tblCU']['CalePlanSituatieCU'].strip('"'),
        y['tblCU']['CaleMemoriuTehnicSS'].strip('"'),
        y['tblCU']['CaleActeFacturare'].strip('"'),
    ]

    x.merge_pdfs_print(pdf_list, path_document_final)

    if os.path.exists(cerere_pdf_path):
        os.remove(cerere_pdf_path)

    x.copy_file(y['tblCU']['CalePlanSituatiePDF'], path_final, director_final, 'Plan situatie.pdf')
    x.copy_file("G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/BC/09.Aviz MApN/Citeste-ma.docx",
                path_final, director_final, 'Citeste-ma.docx')

    print("\nAvizul MApN a fost creat \n")



def aviz_RAJA_Onesti(id_lucrare, path_final):
    director_final = '06.Aviz RAJA Onești'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/BC/06. Aviz RAJA/'f"01.Cerere RAJA{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

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
        'nr_cu': y['tblCU']['NrCU'],
        'data_cu': x.get_date(y['tblCU']['DataCU']),
        'emitent_cu': y['EmitentCU']['denumire_institutie'],
        # Data
        'data': y['astazi'],
    }

    cerere_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))



    # 02. Document de identitate    # punem CUI-ul ROGOTEHNIC

    x.copy_file(y['firma_proiectare']['CaleCertificat'], path_final,
                director_final, '02. Document de identitate.pdf')

    
    # 03. Documente de proprietate

    lista_extrase = []

    with os.scandir(y['tblCU']['CaleExtraseCF'].strip('"')) as entries:
        for entry in entries:
            if entry.is_file() and "Extras" in str(entry):
                lista_extrase.append(entry.path)
    
    path_extrase = os.path.join(
        path_final, director_final, "03. Documente de proprietate.pdf")
    
    x.merge_pdfs_print(lista_extrase, path_extrase)

    # 04. Plan cadastral - punem tot extrasul de CF

    lista_extrase = []

    with os.scandir(y['tblCU']['CaleExtraseCF'].strip('"')) as entries:
        for entry in entries:
            if entry.is_file() and "Extras" in str(entry):
                lista_extrase.append(entry.path)
    
    path_extrase = os.path.join(
        path_final, director_final, "04. Plan cadastral.pdf")
    
    x.merge_pdfs_print(lista_extrase, path_extrase)

    # 05. Extrase CF

    lista_extrase = []

    with os.scandir(y['ExtraseCU']['CaleExtraseCF'].strip('"')) as entries:
        for entry in entries:
            if entry.is_file() and "Extras" in str(entry):
                lista_extrase.append(entry.path)
    
    path_extrase = os.path.join(
        path_final, director_final, "05. Extrase CF.pdf")
    
    x.merge_pdfs_print(lista_extrase, path_extrase)


    # 06. Act de intabulare

    lista_extrase = []

    with os.scandir(y['ExtraseCU']['CaleExtraseCF'].strip('"')) as entries:
        for entry in entries:
            if entry.is_file() and "Extras" in str(entry):
                lista_extrase.append(entry.path)
    
    path_extrase = os.path.join(
        path_final, director_final, "06. Act de intabulare.pdf")
    
    x.merge_pdfs_print(lista_extrase, path_extrase)

    # 07. Plan de situatie
    x.copy_file(y['tblCU']['CalePlanSituatiePDF'], path_final,
                director_final, '07.Plan situatie.pdf')

    # 08. Certificat de urbanism
    x.copy_file(y['tblCU']['CaleCU'], path_final, director_final, '08.Certificat de urbanism.pdf')

    # 09. Memoriu tehnic
    x.copy_file(y['tblCU']['CaleMemoriuTehnicSS'], path_final, director_final, '09.Memoriu Tehnic.pdf')

    # Plan situatie anexa CU    
    x.copy_file(y['tblCU']['CalePlanSituatieCU'], path_final, director_final, '10.Plan situatie anexa CU.pdf')

    # Plan incadrare anexa CU    
    x.copy_file(y['tblCU']['CalePlanIncadrareCU'], path_final, director_final, '11.Plan incadrare anexa CU.pdf')

    # Plan ATR (avizul tehnic de racordare)    
    x.copy_file(y['tblCU']['CaleMemoriuTehnicSS'], path_final, director_final, '12.ATR.pdf')



    # -------------------------------------------------------------------------------------------------------------------------------------------------------------------

    # creez EMAILUL

    x.copy_file("G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/BC/06. Aviz RAJA/Citeste-ma.docx",
                path_final, director_final, 'Citeste-ma.docx')

    print("\nAvizul RAJA-Bacău a fost creat \n")


def aviz_Romprest(id_lucrare, path_final):
    director_final = '10.Aviz Romprest - Bacau'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/BC/10.Aviz Romprest Bacau/'f"Cerere Romprest{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

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
        'nr_cu': y['tblCU']['NrCU'],
        'data_cu': x.get_date(y['tblCU']['DataCU']),
        'emitent_cu': y['EmitentCU']['denumire_institutie'],
        # Data
        'data': y['astazi'],
    }

    cerere_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz Romprest - {y['client']['nume']} conform CU {y['tblCU']['NrCU']}.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['tblCU']['CaleCU'].strip('"'),
        y['tblCU']['CalePlanIncadrareCU'].strip('"'),
        y['tblCU']['CalePlanSituatieCU'].strip('"'),
        y['tblCU']['CaleMemoriuTehnicSS'].strip('"'),
        y['tblCU']['CaleActeFacturare'].strip('"'),
    ]

    x.merge_pdfs(pdf_list, path_document_final)

    # PUNEM PLANUL DE SITUATIE DETALIAT
    x.copy_file(y['tblCU']['CalePlanSituatiePDF'], path_final,
                director_final, 'Plan situatie.pdf')


    if os.path.exists(cerere_pdf_path):
        os.remove(cerere_pdf_path)

    # -----------------------------------------------------------------------------------------------

    # creez EMAILUL

    model_email = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\BC\10.Aviz Romprest Bacau\Model email.docx")

    facturare = x.facturare(id_lucrare)

    context_email = {
        'nume_client': y['client']['nume'],
        'nr_cu': y['tblCU']['NrCU'],
        'data_cu': x.get_date(y['tblCU']['DataCU']),
        'emitent_cu': y['EmitentCU']['denumire_institutie'],
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

    print("\nAvizul Salubritate Romprest - Bacau a fost creat \n")


def acord_Birou_Tehnic_Onesti(id_lucrare, path_final):
    director_final = '11.Acord Birou Tehnic Onesti'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/BC/11.Acord Birou Tehnic Onesti/'f"Cerere Acord Birou Tehnic Onesti{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

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
        'localitate_repr': y['firma_proiectare']['localitate_repr'],
        'adresa_repr': y['firma_proiectare']['adresa_repr'],
        'judet_repr': y['firma_proiectare']['judet_repr'],
        'seria_CI': y['firma_proiectare']['seria_CI'],
        'nr_CI': y['firma_proiectare']['nr_CI'],
        'cnp_repr': y['firma_proiectare']['cnp_repr'],
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
        'nr_cu': y['tblCU']['NrCU'],
        'data_cu': x.get_date(y['tblCU']['DataCU']),
        'emitent_cu': y['EmitentCU']['denumire_institutie'],
        # Data
        'data': y['astazi'],
    }

    cerere_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie Acord de principiu - {y['client']['nume']} conform CU {y['tblCU']['NrCU']}.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['tblCU']['CaleCU'].strip('"'),
        y['tblCU']['CalePlanIncadrareCU'].strip('"'),
        y['tblCU']['CalePlanSituatieCU'].strip('"'),
        y['tblCU']['CaleMemoriuTehnicSS'].strip('"'),
        y['tblCU']['CaleActeFacturare'].strip('"'),
        y['firma_proiectare']['caleCI'].strip('"'),
    ]

    x.merge_pdfs(pdf_list, path_document_final)

    # PUNEM PLANUL DE SITUATIE DETALIAT
    x.copy_file(y['tblCU']['CalePlanSituatiePDF'], path_final, director_final, 'Plan situatie.pdf')


    if os.path.exists(cerere_pdf_path):
        os.remove(cerere_pdf_path)

    # -----------------------------------------------------------------------------------------------

    # creez EMAILUL

    model_email = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\BC\11.Acord Birou Tehnic Onesti\Model email.docx")

    
    context_email = {
        'nume_client': y['client']['nume'],
        'nr_cu': y['tblCU']['NrCU'],
        'data_cu': x.get_date(y['tblCU']['DataCU']),
        'emitent_cu': y['EmitentCU']['denumire_institutie'],
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nume_client': y['client']['nume'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
    }

    x.create_email(model_email, context_email, y['final_destination'])

    print("\nAcord Birou Tehnic - Onești a fost creat \n")



def aviz_ISU_Bacau(id_lucrare, path_final):
    director_final = '12.Aviz ISU - Bacau'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/BC/12.Aviz ISU/'f"01.Cerere aviz ISU{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

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
        'nr_cu': y['tblCU']['NrCU'],
        'data_cu': x.get_date(y['tblCU']['DataCU']),
        'emitent_cu': y['EmitentCU']['denumire_institutie'],
        # Data
        'data': y['astazi'],
    }

    cerere_ISU_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'])

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz ISU {y['client']['nume']} conform CU {y['tblCU']['NrCU']} din {x.get_date(y['tblCU']['DataCU'])}.pdf")

    # -----------------------------------------------------------------------------------------------------------------------------------------

    # creez OPISul
    model_opis = r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\BC\12.Aviz ISU\01.Opis documente.docx"

    date = x.count_pages_ISU(cerere_ISU_pdf_path, y['tblCU']['CaleCU'].strip('"'), y['tblCU']['CalePlanIncadrareCU'].strip(
        '"'), y['tblCU']['CalePlanSituatieCU'].strip('"'), y['tblCU']['CaleMemoriuTehnicSS'].strip('"'), y['tblCU']['CaleActeFacturare'].strip('"'))
    
    context_opis = {
        # lucrare
        'nr_cu': y['tblCU']['NrCU'],
        'data_cu': x.get_date(y['tblCU']['DataCU']),
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
        y['tblCU']['CaleCU'].strip('"'),
        y['tblCU']['CalePlanIncadrareCU'].strip('"'),
        y['tblCU']['CalePlanSituatieCU'].strip('"'),
        y['tblCU']['CaleMemoriuTehnicSS'].strip('"'),
        y['tblCU']['CaleActeFacturare'].strip('"'),
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
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\BC\12.Aviz ISU\Model email.docx")

    context_email = {
        'nume_client': y['client']['nume'],
        'nr_cu': y['tblCU']['NrCU'],
        'data_cu': x.get_date(y['tblCU']['DataCU']),
        'emitent_cu': y['EmitentCU']['denumire_institutie'],
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nume_client': y['client']['nume'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
    }

    x.create_email(model_email, context_email, y['final_destination'])

    print("\nAvizul ISU - Bacau a fost creat \n")



def aviz_OAR(id_lucrare, path_final):
    director_final = '13.Aviz OAR - Bacau'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\BC\13.Aviz OAR\Cerere OAR.docx")

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
        'nr_cu': y['tblCU']['NrCU'],
        'data_cu': x.get_date(y['tblCU']['DataCU']),
        'emitent_cu': y['EmitentCU']['denumire_institutie'],
        # Data
        'data': y['astazi'],
    }



    cerere_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie Aviz OAR - {y['beneficiar']['nume']} conform CU {y['tblCU']['NrCU']} din {x.get_date(y['tblCU']['DataCU'])}.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['tblCU']['CaleCU'].strip('"'),
        y['tblCU']['CalePlanIncadrareCU'].strip('"'),
        y['tblCU']['CalePlanSituatieCU'].strip('"'),
        y['tblCU']['CaleMemoriuTehnicSS'].strip('"'),
    ]

    x.merge_pdfs(pdf_list, path_document_final)

    if os.path.exists(cerere_pdf_path):
        os.remove(cerere_pdf_path)

    # PUNEM PLANUL DE SITUATIE DETALIAT
    x.copy_file(y['tblCU']['CalePlanSituatiePDF'], path_final,
                director_final, 'Plan situatie.pdf')


    # -----------------------------------------------------------------------------------------------

    # creez EMAILUL

    model_email = (r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\BC\13.Aviz OAR\Model email.docx")
    

    context_email = {
        'nume_client': y['client']['nume'],
        'nr_cu': y['tblCU']['NrCU'],
        'data_cu': x.get_date(y['tblCU']['DataCU']),
        'emitent_cu': y['EmitentCU']['denumire_institutie'],
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nume_client': y['client']['nume'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
        

    }

    x.create_email(model_email, context_email, y['final_destination'])

    print("\nAvizul OAR - Bacau a fost creat \n")



def aviz_ANANP_ST_Bacau(id_lucrare, path_final):
    director_final = '01.Aviz ANANP ST-Bacau'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\BC\14.Aviz Arii Naturale\Cerere ANANP ST-Bacau.docx")

    context_Cerere = {
        'nume_firma_proiectare': y['firma_proiectare']['nume'],
        'reprezentant_firma_proiectare': y['firma_proiectare']['reprezentant'],
        'localitate_reprezentant': y['firma_proiectare']['localitate_repr'],
        'adresa_reprezentant': y['firma_proiectare']['adresa_repr'],
        'judet_reprezentant': y['firma_proiectare']['judet_repr'],
        'seria_CI': y['firma_proiectare']['seria_CI'],
        'nr_CI': y['firma_proiectare']['nr_CI'],

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
        model_cerere, context_Cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    # -----------------------------------------------------------------------------------------------

    # creez NOTIFICAREA

    model_Notificare_ANANP = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\BC\14.Aviz Arii Naturale\Notificare.docx")

    context_Notificare_ANANP = {
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
        'descrierea_proiectului': y['tblCU']['descrierea_proiectului'],
        'intocmit': y['intocmit'],
        'verificat': y['verificat'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
    }

    notificare_pdf_path = x.create_document(
        model_Notificare_ANANP, context_Notificare_ANANP, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    # -----------------------------------------------------------------------------------------------

    # creez EMAILUL

    model_Email = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\BC\14.Aviz Arii Naturale\Model email.docx")
    
    facturare = x.facturare(id_lucrare)

    context_Email_ANANP = {
        'nume_client': y['client']['nume'],
        'nr_cu': y['tblCU']['NrCU'],
        'data_cu': x.get_date(y['tblCU']['DataCU']),
        'emitent_cu': y['EmitentCU']['denumire_institutie'],
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

    x.create_email(model_Email, context_Email_ANANP,
                   y['final_destination'])

    # -----------------------------------------------------------------------------------------------

    # creez DOCUMENTUL FINAL

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz ANANP ST-Bacau.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['tblCU']['CaleCU'].strip('"'),
        y['tblCU']['CalePlanIncadrareCU'].strip('"'),
        y['tblCU']['CalePlanSituatieCU'].strip('"'),
        notificare_pdf_path,
        y['tblCU']['CaleActeFacturare'].strip('"'),
    ]

    x.merge_pdfs_print(pdf_list, path_document_final)

    # PUNEM PLANUL DE SITUATIE DETALIAT
    x.copy_file(y['tblCU']['CalePlanSituatiePDF'], path_final,
                director_final, 'Plan situatie.pdf')
    x.copy_file(y['tblCU']['CalePlanSituatieDWG'], path_final,
                director_final, 'Plan situatie.dwg')

    # -----------------------------------------------------------------------------------------------

    # fac curat

    if os.path.exists(cerere_pdf_path):
        os.remove(cerere_pdf_path)

    if os.path.exists(notificare_pdf_path):
        os.remove(notificare_pdf_path)

    print("\nAvizul ANANP ST-Bacău a fost creat \n")