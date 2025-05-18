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
        'descrierea_proiectului': y['tblCU']['DescriereaProiectului'],
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

    x.create_email(model_Email_APM, context_Email_APM_Iasi,
                   y['final_destination'])

    # -----------------------------------------------------------------------------------------------

    # creez DOCUMENTUL FINAL

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz APM Botoșani - {y['client']['nume']} conform CU {y['tblCU']['NrCU']} din {x.get_date(y['tblCU']['DataCU'])}.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['tblCU']['CaleChitantaAPM'].strip('"'),
        y['tblCU']['CaleCU'].strip('"'),
        y['tblCU']['CalePlanIncadrareCU'].strip('"'),
        y['tblCU']['CalePlanSituatieCU'].strip('"'),
        notificare_pdf_path,
        y['tblCU']['CaleActeFacturare'].strip('"'),
    ]

    x.merge_pdfs(pdf_list, path_document_final)


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

        'nr_cu': y['tblCU']['NrCU'],
        'data_cu': x.get_date(y['tblCU']['DataCU']),
        'emitent_cu': y['EmitentCU']['denumire_institutie'],
        # Data
        'data': y['astazi'],
    }

    cerere_EE_pdf_path = x.create_document(
        model_cerere_EE_Delgaz, context_cerere_EE_Delgaz, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz EE - Delgaz - {y['client']['nume']} conform CU {y['tblCU']['NrCU']} din {x.get_date(y['tblCU']['DataCU'])}.pdf")

    pdf_list = [
        cerere_EE_pdf_path,
        y['tblCU']['CaleCU'].strip('"'),
        y['tblCU']['CalePlanIncadrareCU'].strip('"'),
        y['tblCU']['CalePlanSituatieCU'].strip('"'),
        y['tblCU']['CaleMemoriuTehnicSS'].strip('"'),
        y['tblCU']['CaleActeFacturare'].strip('"'),
    ]

    if y['lucrare']['IDClient'] != 1:
        pdf_list.insert(-1, y['tblCU']['CaleActeBeneficiar'].strip('"'))

    x.merge_pdfs(pdf_list, path_document_final)

    # PUNEM PLANUL DE SITUATIE DETALIAT
    x.copy_file(y['tblCU']['CalePlanSituatiePDF'], path_final,
                director_final, 'Plan situatie.pdf')
    x.copy_file(y['tblCU']['CalePlanSituatieDWG'], path_final,
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
        'nr_cu': y['tblCU']['NrCU'],
        'data_cu': x.get_date(y['tblCU']['DataCU']),
        'emitent_cu': y['EmitentCU']['denumire_institutie'],
        # Data
        'data': y['astazi'],
    }

    cerere_GN_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz GN - Delgaz - {y['client']['nume']} conform CU nr. {y['tblCU']['NrCU']} din {x.get_date(y['tblCU']['DataCU'])}.pdf")

    pdf_list = [
        cerere_GN_pdf_path,
        y['tblCU']['CaleCU'].strip('"'),
        y['tblCU']['CalePlanIncadrareCU'].strip('"'),
        y['tblCU']['CalePlanSituatieCU'].strip('"'),
        y['tblCU']['CaleMemoriuTehnicSS'].strip('"'),
        y['tblCU']['CaleActeFacturare'].strip('"'),
    ]

    if y['lucrare']['IDClient'] != 1:
        pdf_list.insert(-1, y['tblCU']['CaleActeBeneficiar'].strip('"'))

    x.merge_pdfs(pdf_list, path_document_final)


        # PUNEM PLANUL DE SITUATIE DETALIAT
    x.copy_file(y['tblCU']['CalePlanSituatiePDF'], path_final,
                director_final, 'Plan situatie.pdf')
    x.copy_file(y['tblCU']['CalePlanSituatieDWG'], path_final,
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
        'nr_cu': y['tblCU']['NrCU'],
        'data_cu': x.get_date(y['tblCU']['DataCU']),
        'emitent_cu': y['EmitentCU']['denumire_institutie'],
        # Data
        'data': y['astazi'],
    }

    cerere_Orange_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))
    os.rename(cerere_Orange_pdf_path, os.path.join(
        path_final, director_final, '01.Cerere.pdf'))

    x.copy_file(y['tblCU']['CaleCU'], path_final,
                director_final, '02.Certificat de urbanism.pdf')
    x.copy_file(y['tblCU']['CalePlanIncadrareCU'], path_final,
                director_final, '03.Plan incadrare in zona.pdf')
    x.copy_file(y['tblCU']['CalePlanSituatieCU'], path_final,
                director_final, '04.Plan situatie.pdf')
    x.copy_file(y['tblCU']['CaleMemoriuTehnicSS'], path_final,
                director_final, '05.Memoriu tehnic.pdf')
    x.copy_file(y['tblCU']['CaleActeFacturare'], path_final,
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
        path_final, director_final, f"Documentatie aviz HCL.pdf")

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
    x.copy_file(y['tblCU']['CalePlanSituatiePDF'], path_final,
                director_final, 'Plan situatie.pdf')
    x.copy_file(y['tblCU']['CalePlanSituatieDWG'], path_final,
                director_final, 'Plan situatie.dwg')

    if os.path.exists(cerere_HCL_pdf_path):
        os.remove(cerere_HCL_pdf_path)

    print("\nAvizul HCL a fost creat \n")


def aviz_SGA(id_lucrare, path_final):
    director_final = '05.Aviz SGA - Botosani'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\BT\05. Aviz SGA\Cerere SGA Botosani.docx")

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
        path_final, director_final, f"Documentatie aviz SGA - Botoșani - {y['client']['nume']} conform CU nr. {y['tblCU']['NrCU']} din {x.get_date(y['tblCU']['DataCU'])}.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['tblCU']['CaleCU'].strip('"'),
        y['tblCU']['CalePlanIncadrareCU'].strip('"'),
        y['tblCU']['CalePlanSituatieCU'].strip('"'),
        y['tblCU']['CaleMemoriuTehnicSS'].strip('"'),
        y['tblCU']['CaleActeFacturare'].strip('"'),
    ]


    x.merge_pdfs_print(pdf_list, path_document_final)

    if os.path.exists(cerere_pdf_path):
        os.remove(cerere_pdf_path)

    # PUNEM PLANUL DE SITUATIE DETALIAT
    x.copy_file(y['tblCU']['CalePlanSituatiePDF'], path_final,
                director_final, 'Plan situatie.pdf')
    x.copy_file(y['tblCU']['CalePlanSituatieDWG'], path_final,
                director_final, 'Plan situatie.dwg')

    # -----------------------------------------------------------------------------------------------
    
    # aici copii fisierul citeste-ma.docx
    x.copy_file("G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/BT/05. Aviz SGA/Citeste-ma.docx",
                path_final, director_final, 'Citeste-ma.docx')

    

    print("\nAvizul SGA - Botosani a fost creat \n")


def aviz_OAR(id_lucrare, path_final):
    director_final = '08.Aviz OAR - Botosani'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\BT\08.Aviz OAR\Model email.docx")

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

    model_email = (r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\BT\08.Aviz OAR\Cerere OAR.docx")
    

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

    print("\nAvizul OAR - Botoșani a fost creat \n")


def aviz_Nova_Apaserv(id_lucrare, path_final):
    director_final = '10.Aviz Nova Apaserv'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/BT/10.Aviz Nova Apaserv/'f"Cerere Nova Apaserv{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")
    
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

        # beneficiar
        'nume_beneficiar': y['beneficiar']['nume'],
        # lucrare
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],

        # Data
        'data': y['astazi'],
    }

    cerere_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie Aviz Nova Apaserv - {y['client']['nume']} conform CU {y['tblCU']['NrCU']} din {x.get_date(y['tblCU']['DataCU'])}.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['tblCU']['CaleCU'].strip('"'),
        y['tblCU']['CalePlanIncadrareCU'].strip('"'),
        y['tblCU']['CalePlanSituatieCU'].strip('"'),
        y['tblCU']['CaleMemoriuTehnicSS'].strip('"'),
        y['tblCU']['CaleActeFacturare'].strip('"'),
    ]

    if y['lucrare']['IDClient'] != 1:
        pdf_list.insert(-1, y['tblCU']['CaleActeBeneficiar'].strip('"'))

    x.merge_pdfs(pdf_list, path_document_final)

    # PUNEM PLANUL DE SITUATIE DETALIAT

    x.copy_file(y['tblCU']['CalePlanSituatiePDF'], path_final,
                director_final, 'Plan situatie.pdf')
    x.copy_file(y['tblCU']['CalePlanSituatieDWG'], path_final,
                director_final, 'Plan situatie.dwg')

    if os.path.exists(cerere_pdf_path):
        os.remove(cerere_pdf_path)

    # -----------------------------------------------------------------------------------------------

    # creez EMAILUL

    model_email = (r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\BT\10.Aviz Nova Apaserv\Model email.docx")

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

    print("\n Avizul Nova Apaserv a fost creat \n")


def aviz_Cultura(id_lucrare, path_final):
    director_final = '11.Aviz Cultura - Botosani'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/MODELE/BT/11.Aviz Cultura Botosani/'f"Cerere Cultura{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

    calcul = float(y['tblCU']['SuprafataOcupata']) * 3.00
    total_aviz = round(calcul, 2)

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
        'sup_mp': y['tblCU']['SuprafataOcupata'],
        'total_aviz': total_aviz,
        # Data
        'data': y['astazi'],
    }

    cerere_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie aviz Cultura - {y['client']['nume']} conform CU {y['tblCU']['NrCU']} din {x.get_date(y['tblCU']['DataCU'])}.pdf")

    path_document_printabil = os.path.join(
        path_final, director_final, f"Documentatie aviz Cultura - DE PRINTAT.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['tblCU']['CaleCU'].strip('"'),
        y['tblCU']['CalePlanIncadrareCU'].strip('"'),
        y['tblCU']['CalePlanSituatieCU'].strip('"'),
        y['tblCU']['CaleMemoriuTehnicSS'].strip('"'),
        y['tblCU']['CaleActeFacturare'].strip('"'),
    ]

    x.merge_pdfs(pdf_list, path_document_final)
    x.merge_pdfs_print(pdf_list, path_document_printabil)

    if os.path.exists(cerere_pdf_path):
        os.remove(cerere_pdf_path)
    
    # -----------------------------------------------------------------------------------------------

    # creez EMAILUL

    model_email = (r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\BT\11.Aviz Cultura Botosani\Model email.docx")

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

    print("\n Avizul Cultura Botoșani a fost creat \n")


def aviz_Biroul_Rutier(id_lucrare, path_final):
    director_final = '09.Aviz Biroul Rutier'
    y = x.get_data(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\BT\09.Aviz Politia Rutiera\Cerere Politia.docx")
    
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

        # beneficiar
        'nume_beneficiar': y['beneficiar']['nume'],
        # lucrare
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],

        # Data
        'data': y['astazi'],
    }

    cerere_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], y['firma_proiectare']['CaleStampila'].strip('"'))

    path_document_final = os.path.join(
        path_final, director_final, f"Documentatie Aviz Principiu Politia Rutiera - {y['client']['nume']} conform CU {y['tblCU']['NrCU']} din {x.get_date(y['tblCU']['DataCU'])}.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['tblCU']['CaleChitantaPolitie'].strip('"'),
        y['tblCU']['CaleCU'].strip('"'),
        y['tblCU']['CalePlanIncadrareCU'].strip('"'),
        y['tblCU']['CalePlanSituatieCU'].strip('"'),
        y['tblCU']['CaleMemoriuTehnicSS'].strip('"'),
        y['tblCU']['CaleActeFacturare'].strip('"'),
        y['tblCU']['CaleActeBeneficiar'].strip('"'),
    ]


    x.merge_pdfs(pdf_list, path_document_final)


    if os.path.exists(cerere_pdf_path):
        os.remove(cerere_pdf_path)

    # -----------------------------------------------------------------------------------------------

    # creez EMAILUL

    model_email = (r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\BT\09.Aviz Politia Rutiera\Model email.docx")


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

    print("\n Avizul Biroul Rutier Botoșani a fost creat \n")