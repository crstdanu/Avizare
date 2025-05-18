import os
import functii as x
from dateutil.relativedelta import relativedelta



# ----------------   INCEPERE lucrari   ---------------------------------------------------------


def conventie_lucrari(id_lucrare, path_final):
    director_final = '01.Conventie lucrari'
    cale_stampila = "G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/DOCUMENTE/Stampila - RGT.png"
    y = x.get_data_executie(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\Executie\01. Pentru incepere\Conventie lucrari\Cerere CL.docx")

    context_cerere = {
        'nr_cl': y['tblIncepereExecutie']['NrCL'],
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nume_client': y['client']['nume'],
        # Data
        'data': y['astazi'],
    }

    cerere_pdf_path = x.create_document(
        model_cerere, context_cerere, y['final_destination'], cale_stampila)

    x.move_file(cerere_pdf_path, path_final,
                director_final, f"01.Cerere CL.pdf")

    x.copy_file(y['tblIncepereExecutie']['CaleInstruireColectiva'], path_final,
                director_final, '02.SCAN - Instruire colectiva.pdf')

    x.copy_file_prefix(y['tblIncepereExecutie']['CaleContractRacordare'],
                       path_final, director_final, '03.')

    x.copy_file_prefix(y['tblIncepereExecutie']['CaleContractExecutie'],           
                       path_final, director_final, '03.')

    if y['tblCU']['CaleAvizCTE']:
        x.copy_file_prefix(y['tblCU']['CaleAvizCTE'], path_final, director_final, '04.')

    if y['tblCU']['CaleATR']:
        x.copy_file_prefix(y['tblCU']['CaleATR'], path_final, director_final, '04.')

    x.copy_file(y['tblIncepereExecutie']['CaleMemoriuTehnicACScanat'], path_final,
                director_final, '05. Memoriu tehnic PTH.pdf')

    x.copy_file(y['tblIncepereExecutie']['CalePlanIncadrarePTH'], path_final,
                director_final, '06. Plan incadrare PTH.pdf')

    x.copy_file(y['tblIncepereExecutie']['CalePlanSituatiePTH'], path_final,
                director_final, '07. Plan situatie PTH.pdf')

    x.copy_file(y['tblIncepereExecutie']['CaleSchemaMonofilaraJT'], path_final,
                director_final, '08. Schema monofilara JT.pdf')

    if y['tblIncepereExecutie']['CaleSchemaMonofilaraMT']:
        x.copy_file(y['tblIncepereExecutie']['CaleSchemaMonofilaraMT'],
                    path_final, director_final, '09. Schema monofilara MT.pdf')

    path_document_final = os.path.join(
        path_final, director_final, f"10. AC+planse.pdf")

    pdf_list = [
        y['tblIncepereExecutie']['CaleACScanat'].strip('"'),
        y['tblIncepereExecutie']['CalePlanIncadrareACScanat'].strip('"'),
        y['tblIncepereExecutie']['CalePlanSituatieACScanat'].strip('"'),

    ]

    x.merge_pdfs_print(pdf_list, path_document_final)





    # -----------------------------------------------------------------------------------------------

    # creez EMAILUL

    model_email = (
        'G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/Executie/01. Pentru incepere/Conventie lucrari/'f"Model email{' - Iasi' if y['tblCU']['EmitentCU'] == 1 else ''}.docx")

    context_email = {
        'nume_client': y['client']['nume'],
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nume_client': y['client']['nume'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
        'nr_cl': y['tblIncepereExecutie']['NrCL'],
        'data': y['astazi'],
    }

    x.create_email(model_email, context_email, y['final_destination'])

    print("\nConvenția de lucrări a fost creată \n")


def anunt_UAT_incepere (id_lucrare, path_final):
    director_final = '02. Anunt UAT'
    cale_stampila = "G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/DOCUMENTE/Stampila - RGT.png"
    y = x.get_data_executie(path_final, director_final, id_lucrare)

    # -----------------------------------------------------------------------------------------------

    # creez CEREREA

    model_cerere = (
        r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\Executie\01. Pentru incepere\Anunt UAT\Cerere Anunt UAT.docx")

    context_cerere = {
        'nume_uat': y['EmitentAC']['denumire_institutie'],
        'localitate_uat': y['EmitentAC']['localitate'],
        'adresa_uat': y['EmitentAC']['adresa'],
        'judet_uat': y['EmitentAC']['judet'],

        'nume_firma_executie': y['firma_executie']['nume'],
        'localitate_firma_executie': y['firma_executie']['localitate'],
        'adresa_firma_executie': y['firma_executie']['adresa'],
        'judet_firma_executie': y['firma_executie']['judet'],
        'email_firma_executie': y['firma_executie']['email'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],

        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nr_ac': y['tblIncepereExecutie']['NumarAC'],
        'data_ac': x.get_date(y['tblIncepereExecutie']['DataAC']),
        'data_incepere': x.get_date(y['tblIncepereExecutie']['DataIncepereExecutie']),

        'nume_client': y['client']['nume'],

        # Data
        'data': y['astazi'],
    }

    cerere_pdf_path = x.create_document(model_cerere, context_cerere, y['final_destination'], cale_stampila)


    path_document_final = os.path.join(path_final, director_final, f"Documentatie Anunt UAT - {y['client']['nume']}.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['tblIncepereExecutie']['CaleACScanat'].strip('"'),
        y['tblIncepereExecutie']['CalePlanIncadrareACScanat'].strip('"'),
        y['tblIncepereExecutie']['CalePlanSituatieACScanat'].strip('"'),
    ]

    x.merge_pdfs_print(pdf_list, path_document_final)

    print("\nAnunțul UAT a fost creat \n")



def grafic_executie(id_lucrare, path_final):
    director_final = '03. Grafic executie'
    cale_stampila = "G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/DOCUMENTE/Stampila - RGT.png"
    y = x.get_data_executie(path_final, director_final, id_lucrare)

    

    # -------------------------------------------------------------------------------------------------------------------

    # Creez Graficul
    if y['lucrare']['localitate'] == "Municipiul Iași":
        file_path = "G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/Executie/01. Pentru incepere/Grafic executie/MODEL - Iasi.xlsx"
    else:
        file_path = "G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/Executie/01. Pentru incepere/Grafic executie/MODEL.xlsx"

    logo_path = "G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/DOCUMENTE/Logo RGT.png"  # Specify the path to your image file
    stampila_path = "G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/DOCUMENTE/Stampila - RGT.png"  # Specify the path to your image file
    final_destination = "G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/Executie/01. Pentru incepere/Grafic executie/output.xlsx"


    an_lucrare = x.get_year(y['tblIncepereExecutie']['DataIncepereGrafic'])
    data_incepere = x.get_date(y['tblIncepereExecutie']['DataIncepereGrafic'])
    data_finalizare = x.get_date(y['tblIncepereExecutie']['DataFinalizareGrafic'])
    luna_incepere = x.get_month(y['tblIncepereExecutie']['DataIncepereGrafic'])
    luna_finalizare = x.get_month(y['tblIncepereExecutie']['DataFinalizareGrafic'])

    context = {
        'file_path': file_path,
        'logo_path': logo_path,
        'stampila_path': stampila_path,

        'sapare_sant': y['tblIncepereExecutie']['DPSapareSant'],
        'pozare_cablu': y['tblIncepereExecutie']['DPPozareCablu'],
        'montare_ptav': y['tblIncepereExecutie']['DPMontarePTAV'],
        'montare_stalpi': y['tblIncepereExecutie']['DPMontareStalpi'],
        'executie_foraj': y['tblIncepereExecutie']['DPExecutieForaj'],
        'montare_firide': y['tblIncepereExecutie']['DPMontareFiride'],
        'aducere_stare_initiala': y['tblIncepereExecutie']['DPStareInitiala'],

        'emitent_ac': y['EmitentAC']['denumire_institutie'],
        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nume_client': y['client']['nume'],
        'an_lucrare': an_lucrare,
        'data_incepere': data_incepere,
        'luna_incepere': luna_incepere,
        'data_finalizare': data_finalizare,
        'luna_finalizare': luna_finalizare,
    }


    path_document_final = os.path.join(path_final, director_final, "Grafic executie.xlsx")

    if y['lucrare']['localitate'] == "Municipiul Iași":
        x.genereaza_grafic_executie_iasi(context, path_document_final)
    else:
        x.genereaza_grafic_executie(context, path_document_final)

    path_grafic_executie = x.xlsx_to_pdf(path_document_final)

    # sterg fisierul excel
    if os.path.exists(path_document_final):
        os.remove(path_document_final)
    
    # Am terminat graficul de executie

    # -------------------------------------------------------------------------------------------------------------------
    # Creez cererea

    if y['tblCU']['EmitentCU'] == 1:
        model_cerere = ('G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/Executie/01. Pentru incepere/Grafic executie/'f"Cerere Grafic executie Iasi{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")
    else:
        model_cerere = (r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\Executie\01. Pentru incepere\Grafic executie\Cerere Grafic executie - provincie.docx")

    context_cerere = {
        'nume_uat': y['EmitentAC']['denumire_institutie'],
        'localitate_uat': y['EmitentAC']['localitate'],
        'adresa_uat': y['EmitentAC']['adresa'],
        'judet_uat': y['EmitentAC']['judet'],

        'nume_firma_executie': y['firma_executie']['nume'],
        'localitate_firma_executie': y['firma_executie']['localitate'],
        'adresa_firma_executie': y['firma_executie']['adresa'],
        'judet_firma_executie': y['firma_executie']['judet'],
        'judet_firma_executie': y['firma_executie']['judet'],
        'email_firma_executie': y['firma_executie']['email'],
        'cui_firma_executie': y['firma_executie']['CUI'],
        'nr_reg_com': y['firma_executie']['NrRegCom'],
        'reprezentant_firma_executie': y['firma_executie']['reprezentant'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],

        'nume_lucrare': y['lucrare']['nume'],
        'localitate_lucrare': y['lucrare']['localitate'],
        'adresa_lucrare': y['lucrare']['adresa'],
        'judet_lucrare': y['lucrare']['judet'],
        'nr_ac': y['tblIncepereExecutie']['NumarAC'],
        'emitent_ac': y['EmitentAC']['denumire_institutie'].upper(),

        'data_ac': x.get_date(y['tblIncepereExecutie']['DataAC']),
        'data_incepere': x.get_date(y['tblIncepereExecutie']['DataIncepereExecutie']),

        'nume_client': y['client']['nume'],
        'nume_beneficiar': y['beneficiar']['nume'],

        # Data
        'data': y['astazi'],
    }

    cerere_pdf_path = x.create_document(model_cerere, context_cerere, y['final_destination'], cale_stampila)

    path_document_final = os.path.join(path_final, director_final, f"Documentatie Grafic execuție - DE PRINTAT.pdf")

    pdf_list = [
        cerere_pdf_path,
        path_grafic_executie,
        path_grafic_executie,
        y['tblIncepereExecutie']['CaleACScanat'].strip('"'),
        y['tblIncepereExecutie']['CalePlanIncadrareACScanat'].strip('"'),
        y['tblIncepereExecutie']['CalePlanSituatieACScanat'].strip('"'),
        y['tblIncepereExecutie']['CaleMemoriuTehnicACScanat'].strip('"'),
    ]

    if y['tblCU']['EmitentCU'] == 1:
        pdf_list.append(y['tblIncepereExecutie']['CaleContractCitadin'].strip('"'))

    if y['tblIncepereExecutie']['CaleContractSpatiiVerzi'] and y['tblCU']['EmitentCU'] == 1:
        pdf_list.append(y['tblIncepereExecutie']['CaleContractSpatiiVerzi'].strip('"'))

    if y['tblIncepereExecutie']['CaleContractForaj'] and y['tblCU']['EmitentCU'] == 1:
        pdf_list.append(y['tblIncepereExecutie']['CaleContractForaj'].strip('"'))

    x.merge_pdfs_print(pdf_list, path_document_final)
    
    if os.path.exists(path_grafic_executie):
        os.remove(path_grafic_executie)
        
    if os.path.exists(cerere_pdf_path):
        os.remove(cerere_pdf_path)

    print("\nGraficul de lucrari a fost creat\n")



def decizie_numire_personal(id_lucrare, path_final):
    director_final = '04. Decizie atributiuni personal'
    cale_stampila = "G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/DOCUMENTE/Stampila - RGT.png"
    y = x.get_data_executie(path_final, director_final, id_lucrare)

    model_cerere = (r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\Executie\01. Pentru incepere\Decizie atributiuni personal\Decizie personal.docx")

    if y['rte_constructii']['nume']:
        nume_rte_constructii = y['rte_constructii']['nume']
    else:
        nume_rte_constructii = 'Nu este cazul'

    context_cerere = {
        'nr_decizie': y['tblIncepereExecutie']['NrDeciziePersonal'],
        'reprezentant_firma_executie': y['firma_executie']['reprezentant'],
        'nume_firma_executie': y['firma_executie']['nume'],
        'localitate_firma_executie': y['firma_executie']['localitate'],
        'adresa_firma_executie': y['firma_executie']['adresa'],
        'judet_firma_executie': y['firma_executie']['judet'],
        'cui_firma_executie': y['firma_executie']['CUI'],
        'nr_reg_com': y['firma_executie']['NrRegCom'],

        'nume_lucrare': y['lucrare']['nume'],

        'nume_responsabil_ssm': y['responsabil_ssm']['nume'],
        'nume_rte_electric': y['rte']['nume'],
        'nume_rte_constructii': nume_rte_constructii,    

        # Data
        'data': y['astazi'],
    }

    cerere_pdf_path = x.create_document(model_cerere, context_cerere, y['final_destination'], cale_stampila)


    print("\nDecizia atributiuni personal a fost creată\n")




def ordin_incepere(id_lucrare, path_final):
    pass


# ----------------   FINALIZARE lucrari   --------------------------------------------------------

def pentru_referat_DS(id_lucrare, path_final):
    pass

def declaratie_ITL(id_lucrare, path_final):
    director_final = '01. Declaratie ITL'
    cale_stampila = "G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/DOCUMENTE/Stampila - RGT.png"
    y = x.get_data_finalizare(path_final, director_final, id_lucrare)
    valoare_reala = float(y['tblFinalizare']['ValoareReala'])


    model_cerere = ('G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/Executie/03.Pentru finalizare/01. Anunt ITL/'f"ITL-064 - Iasi{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")
 
    
    expirare_executie = x.aduna_luni(y['tblIncepereExecutie']['DataIncepereExecutie'], int(y['tblIncepereExecutie']['ValabilitateExecutie']))
    valoare_ac = float(y['tblIncepereExecutie']['ValoareAC'])
    taxa_ac = x.custom_round(valoare_ac)
    taxa_reala = x.custom_round(valoare_reala * 0.01)
    diferenta_taxa = x.diferenta_taxa(taxa_ac, taxa_reala)
    taxa_reala = x.taxa_reala(valoare_reala, valoare_ac)

    context_cerere = {
        'nume_client': y['client']['nume'],
        'nume_firma_executie': y['firma_executie']['nume'],
        'nume_manager_proiect': y['manager_proiect']['nume'].upper(),
        'nr_ac': y['tblIncepereExecutie']['NumarAC'],
        'data_ac': x.get_date(y['tblIncepereExecutie']['DataAC']),
        'valoare_ac': f"{y['tblIncepereExecutie']['ValoareAC']:.2f}",
        'emitent_ac': y['EmitentAC']['denumire_institutie'],
        'expirare_executie': expirare_executie,
        'valoare_reala': f"{valoare_reala:.2f}",
        'taxa_reala': taxa_reala,
        'taxa_ac': f"{taxa_ac:.2f}", 
        'diferenta_taxa': diferenta_taxa,
        # Data
        'data': y['astazi'],
    }

    cerere_pdf_path = x.create_document(model_cerere, context_cerere, y['final_destination'], cale_stampila)

    path_document_final = os.path.join(path_final, director_final, f"Documentatie ITL - {y['client']['nume']}.pdf")


    pdf_list = [
        cerere_pdf_path,
        y['firma_executie']['CaleCertificat'].strip('"'),
        y['firma_executie']['CaleCI'].strip('"'),

        y['tblIncepereExecutie']['CaleACScanat'].strip('"'),
        y['tblIncepereExecutie']['CalePlanIncadrareACScanat'].strip('"'),
        y['tblIncepereExecutie']['CalePlanSituatieACScanat'].strip('"'),
        y['tblIncepereExecutie']['CaleMemoriuTehnicACScanat'].strip('"'),

        y['tblIncepereExecutie']['CaleContractRacordare'].strip('"'),
        y['tblIncepereExecutie']['CaleContractExecutie'].strip('"'),
        y['tblFinalizare']['CaleDevizFinal'].strip('"'),

        y['tblFinalizare']['CaleReferatDS'].strip('"'),
        y['tblFinalizare']['CaleRaportProiectant'].strip('"'),
        
    ]

    if y['tblFinalizare']['CaleImputernicireDelgaz']:
        pdf_list.insert(1, y['tblFinalizare']['CaleImputernicireDelgaz'].strip('"'))
    ### --------------------------- doar cand avem Dispozitie de santier --------------------------------------------- ###
    if y['tblFinalizare']['CaleDispozitieSantier']:
        pdf_list.insert(-1, y['tblFinalizare']['CaleDispozitieSantier'].strip('"'))
    if y['tblFinalizare']['CalePlanIncadrareDS']:
        pdf_list.insert(-1, y['tblFinalizare']['CalePlanIncadrareDS'].strip('"'))
    if y['tblFinalizare']['CalePlanSituatieDS']:
        pdf_list.insert(-1, y['tblFinalizare']['CalePlanSituatieDS'].strip('"'))
    if y['tblFinalizare']['CaleRaspunsUatDS']:
        pdf_list.insert(-1, y['tblFinalizare']['CaleRaspunsUatDS'].strip('"'))
    ### ----------------------------------------------------------------------------------------------------------------####

    if y['tblIncepereExecutie']['CaleDovadaPlataAC']:
        pdf_list.append(y['tblIncepereExecutie']['CaleDovadaPlataAC'].strip('"'))
    if y['tblIncepereExecutie']['CaleDovadaPlataISC']:
        pdf_list.append(y['tblIncepereExecutie']['CaleDovadaPlataISC'].strip('"'))

    if y['tblFinalizare']['CaleFacturiRGT']:
        pdf_list.append(y['tblFinalizare']['CaleFacturiRGT'].strip('"'))
    if y['tblFinalizare']['CaleDovadaPlataFacturi']:
        pdf_list.append(y['tblFinalizare']['CaleDovadaPlataFacturi'].strip('"'))

    
    x.merge_pdfs(pdf_list, path_document_final)

    if os.path.exists(cerere_pdf_path):
        os.remove(cerere_pdf_path)


    ### ------------------------------------------------------------- Creez EMAILUL ------------------------------------------------ ###


    model_email = (r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\Executie\03.Pentru finalizare\01. Anunt ITL\Model email.docx")

    context_email = {
        'nume_client': y['client']['nume'],
        'nr_ac': y['tblIncepereExecutie']['NumarAC'],
        'data_ac': x.get_date(y['tblIncepereExecutie']['DataAC']),
        'nume_lucrare': y['lucrare']['nume'],

        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
    }

    x.create_email(model_email, context_email, y['final_destination'])


    print("\n Declaratia ITL a fost creată \n")



def anunt_UAT_finalizare(id_lucrare, path_final):
    director_final = '02. Anunt UAT - Finalizare'
    cale_stampila = "G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/DOCUMENTE/Stampila - RGT.png"
    y = x.get_data_finalizare(path_final, director_final, id_lucrare)


    model_cerere = ('G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/Executie/03.Pentru finalizare/02. Anunt UAT/'f"Anunt UAT{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")

    valoare_reala = float(y['tblFinalizare']['ValoareReala'])  

    context_cerere = {
        'emitent_ac': y['EmitentAC']['denumire_institutie'].upper(),

        'nume_client': y['client']['nume'],

        'nume_firma_executie': y['firma_executie']['nume'],
        'nr_reg_com': y['firma_executie']['NrRegCom'],
        'cui_firma_executie': y['firma_executie']['CUI'],
        'localitate_firma_executie': y['firma_executie']['localitate'],
        'adresa_firma_executie': y['firma_executie']['adresa'],
        'judet_firma_executie': y['firma_executie']['judet'],
        'repr_firma_executie': y['firma_executie']['reprezentant'],
        'serie_ci': y['firma_executie']['seria_CI'],
        'nr_ci': y['firma_executie']['nr_CI'],
        'cnp_repr': y['firma_executie']['cnp_repr'],
        'telefon_contact': y['contact']['telefon'],
        'email_firma_executie': y['firma_executie']['email'],

        'nr_ac': y['tblIncepereExecutie']['NumarAC'],
        'data_ac': x.get_date(y['tblIncepereExecutie']['DataAC']),

        'data_finalizare': x.get_date(y['tblFinalizare']['DataFinalizare']),

        'nume_lucrare': y['lucrare']['nume'],

        'valoare_reala': f"{valoare_reala:.2f}",
        # Data
        'data': y['astazi'],
    }

    cerere_pdf_path = x.create_document(model_cerere, context_cerere, y['final_destination'], cale_stampila)

    path_document_final = os.path.join(path_final, director_final, f"Anunț UAT - Finalizare - {y['client']['nume']}.pdf")


    pdf_list = [
        cerere_pdf_path,
        cerere_pdf_path,
        y['firma_executie']['CaleCertificat'].strip('"'),
        y['firma_executie']['CaleCI'].strip('"'),
        y['tblIncepereExecutie']['CaleACScanat'].strip('"'),
        y['tblIncepereExecutie']['CalePlanIncadrareACScanat'].strip('"'),
        y['tblIncepereExecutie']['CalePlanSituatieACScanat'].strip('"'),
        y['tblIncepereExecutie']['CaleMemoriuTehnicACScanat'].strip('"'),
    ]

    if y['tblFinalizare']['CaleImputernicireDelgaz']:
        pdf_list.insert(2, y['tblFinalizare']['CaleImputernicireDelgaz'].strip('"'))

    if y['tblFinalizare']['CaleAdeverintaISC']:
        pdf_list.append(y['tblFinalizare']['CaleAdeverintaISC'].strip('"'))
    if y['tblFinalizare']['CaleITL']:
        pdf_list.append(y['tblFinalizare']['CaleITL'].strip('"'))
    if y['tblFinalizare']['CaleDovadaRegularizareTaxaAC']:
        pdf_list.append(y['tblFinalizare']['CaleDovadaRegularizareTaxaAC'].strip('"'))
    ### --------------------------- doar cand avem Dispozitie de santier --------------------------------------------- ###
    if y['tblFinalizare']['CaleDispozitieSantier']:
        pdf_list.insert(-1, y['tblFinalizare']['CaleDispozitieSantier'].strip('"'))
    if y['tblFinalizare']['CalePlanIncadrareDS']:
        pdf_list.insert(-1, y['tblFinalizare']['CalePlanIncadrareDS'].strip('"'))
    if y['tblFinalizare']['CalePlanSituatieDS']:
        pdf_list.insert(-1, y['tblFinalizare']['CalePlanSituatieDS'].strip('"'))
    if y['tblFinalizare']['CaleRaspunsUatDS']:
        pdf_list.insert(-1, y['tblFinalizare']['CaleRaspunsUatDS'].strip('"'))
    ### ----------------------------------------------------------------------------------------------------------------####
    if y['tblFinalizare']['CaleReferatDS']:
        pdf_list.append(y['tblFinalizare']['CaleReferatDS'].strip('"'))
    if y['tblFinalizare']['CaleRaportProiectant']:
        pdf_list.append(y['tblFinalizare']['CaleRaportProiectant'].strip('"'))

    x.merge_pdfs_print(pdf_list, path_document_final)

    if os.path.exists(cerere_pdf_path):
        os.remove(cerere_pdf_path)

    print("\n Anuntul UAT Iași a fost creat \n")



def anunt_UAT_declaratie_ITL(id_lucrare, path_final):
    director_final = '01. Anunt UAT + ITL - Finalizare'
    cale_stampila = "G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/DOCUMENTE/Stampila - RGT.png"
    y = x.get_data_finalizare(path_final, director_final, id_lucrare)
    valoare_reala = float(y['tblFinalizare']['ValoareReala'])   # valoarea fara proiectare


    ## -------------------------------------------- Anunt UAT ------------------------------------------ ##

    model_cerere_uat = ('G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/Executie/03.Pentru finalizare/02. Anunt UAT/'f"Anunt UAT{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")
    
    context_cerere_uat = {
        'emitent_ac': y['EmitentAC']['denumire_institutie'].upper(),

        'nume_client': y['client']['nume'],

        'nume_firma_executie': y['firma_executie']['nume'],
        'nr_reg_com': y['firma_executie']['NrRegCom'],
        'cui_firma_executie': y['firma_executie']['CUI'],
        'localitate_firma_executie': y['firma_executie']['localitate'],
        'adresa_firma_executie': y['firma_executie']['adresa'],
        'judet_firma_executie': y['firma_executie']['judet'],
        'repr_firma_executie': y['firma_executie']['reprezentant'],
        'serie_ci': y['firma_executie']['seria_CI'],
        'nr_ci': y['firma_executie']['nr_CI'],
        'cnp_repr': y['firma_executie']['cnp_repr'],
        'telefon_contact': y['contact']['telefon'],
        'email_firma_executie': y['firma_executie']['email'],

        'nr_ac': y['tblIncepereExecutie']['NumarAC'],
        'data_ac': x.get_date(y['tblIncepereExecutie']['DataAC']),

        'data_finalizare': x.get_date(y['tblFinalizare']['DataFinalizare']),

        'nume_lucrare': y['lucrare']['nume'],

        'valoare_reala': f"{valoare_reala:.2f}",
        # Data
        'data': y['astazi'],
    }

    cerere_uat_pdf_path = x.create_document(model_cerere_uat, context_cerere_uat, y['final_destination'], cale_stampila)


    ## --------------------------------------------- declaratie ITL -------------------------------------------------- ##

    
    if y['tblCU']['EmitentCU'] == 2:
        model_declaratie_itl = ('G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/Executie/03.Pentru finalizare/01. Anunt ITL/'f"ITL-064 - Tomesti{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")
    elif y['tblCU']['EmitentCU'] == 3:
        model_declaratie_itl = ('G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/Executie/03.Pentru finalizare/01. Anunt ITL/'f"ITL-064 - Miroslava{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")
    elif y['tblCU']['EmitentCU'] == 7:
        model_declaratie_itl = ('G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/Executie/03.Pentru finalizare/01. Anunt ITL/'f"ITL-064 - Barnova{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")
    elif y['tblCU']['EmitentCU'] == 15:
        model_declaratie_itl = ('G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/Executie/03.Pentru finalizare/01. Anunt ITL/'f"ITL-064 - Vama Suceava{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")
    elif y['tblCU']['EmitentCU'] == 14:
        model_declaratie_itl = ('G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/Executie/03.Pentru finalizare/01. Anunt ITL/'f"ITL-064 - Ciurea{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")
    else:
        print("NU am model de cerere ITL pentru aceastra localitate")


    expirare_executie = x.aduna_luni(y['tblIncepereExecutie']['DataIncepereExecutie'], int(y['tblIncepereExecutie']['ValabilitateExecutie']))
    valoare_ac = float(y['tblIncepereExecutie']['ValoareAC'])
    taxa_ac = x.custom_round(valoare_ac)
    taxa_reala = x.custom_round(valoare_reala * 0.01)
    diferenta_taxa = x.diferenta_taxa(taxa_ac, taxa_reala)
    taxa_reala = x.taxa_reala(valoare_reala, valoare_ac)


    context_declaratie_itl = {
        'nume_client': y['client']['nume'],
        'nume_firma_executie': y['firma_executie']['nume'],
        'nume_manager_proiect': y['manager_proiect']['nume'].upper(),
        'nr_ac': y['tblIncepereExecutie']['NumarAC'],
        'data_ac': x.get_date(y['tblIncepereExecutie']['DataAC']),
        'valoare_ac': f"{y['tblIncepereExecutie']['ValoareAC']:.2f}",
        'emitent_ac': y['EmitentAC']['denumire_institutie'],
        'expirare_executie': expirare_executie,

        'valoare_reala': f"{valoare_reala:.2f}",
        'taxa_reala': taxa_reala,
        'taxa_ac': f"{taxa_ac:.2f}", 
        'diferenta_taxa': diferenta_taxa,
        # Data
        'data': y['astazi'],
    }

    declaratie_itl_pdf_path = x.create_document(model_declaratie_itl, context_declaratie_itl, y['final_destination'], cale_stampila)


    ## ---------------------------------------------------------- document final - de trimis email ------------------------------------------------##

    path_document_final = os.path.join(path_final, director_final, f"Anunț UAT + ITL - Finalizare - {y['client']['nume']}.pdf")

    pdf_list = [
        cerere_uat_pdf_path,
        declaratie_itl_pdf_path,
        y['firma_executie']['CaleCertificat'].strip('"'),
        y['firma_executie']['CaleCI'].strip('"'),
        y['tblIncepereExecutie']['CaleACScanat'].strip('"'),
        y['tblIncepereExecutie']['CalePlanIncadrareACScanat'].strip('"'),
        y['tblIncepereExecutie']['CalePlanSituatieACScanat'].strip('"'),
        y['tblIncepereExecutie']['CaleMemoriuTehnicACScanat'].strip('"'),
        y['tblFinalizare']['CaleDevizFinal'].strip('"'),
    ]

    ### --------------------------- doar cand avem Dispozitie de santier --------------------------------------------- ###
    if y['tblFinalizare']['CaleDispozitieSantier']:
        pdf_list.insert(-1, y['tblFinalizare']['CaleDispozitieSantier'].strip('"'))

    if y['tblFinalizare']['CalePlanIncadrareDS']:
        pdf_list.insert(-1, y['tblFinalizare']['CalePlanIncadrareDS'].strip('"'))

    if y['tblFinalizare']['CalePlanSituatieDS']:
        pdf_list.insert(-1, y['tblFinalizare']['CalePlanSituatieDS'].strip('"'))

    if y['tblFinalizare']['CaleRaspunsUatDS']:
        pdf_list.insert(-1, y['tblFinalizare']['CaleRaspunsUatDS'].strip('"'))
    ### ----------------------------------------------------------------------------------------------------------------####


    if y['tblFinalizare']['CaleImputernicireDelgaz']:
        pdf_list.insert(2, y['tblFinalizare']['CaleImputernicireDelgaz'].strip('"'))

    if y['tblFinalizare']['CaleReferatDS']:
        pdf_list.insert(-1, y['tblFinalizare']['CaleReferatDS'].strip('"'))

    if y['tblFinalizare']['CaleRaportProiectant']:
        pdf_list.insert(-1, y['tblFinalizare']['CaleRaportProiectant'].strip('"'))


    x.merge_pdfs(pdf_list, path_document_final)


    ## --------------------------------------- document final - de printat ---------------------------------------- ##

    path_document_final_print = os.path.join(path_final, director_final, f"Anunț UAT + ITL - Finalizare - de printat.pdf")

    pdf_list_print = [
        cerere_uat_pdf_path,
        declaratie_itl_pdf_path,
        cerere_uat_pdf_path,
        declaratie_itl_pdf_path,
        declaratie_itl_pdf_path,
        y['firma_executie']['CaleCertificat'].strip('"'),
        y['firma_executie']['CaleCI'].strip('"'),
        y['tblIncepereExecutie']['CaleACScanat'].strip('"'),
        y['tblIncepereExecutie']['CalePlanIncadrareACScanat'].strip('"'),
        y['tblIncepereExecutie']['CalePlanSituatieACScanat'].strip('"'),
        y['tblIncepereExecutie']['CaleMemoriuTehnicACScanat'].strip('"'),
        y['tblFinalizare']['CaleDevizFinal'].strip('"'),
    ]

    ### --------------------------- doar cand avem Dispozitie de santier --------------------------------------------- ###
    if y['tblFinalizare']['CaleDispozitieSantier']:
        pdf_list.insert(-1, y['tblFinalizare']['CaleDispozitieSantier'].strip('"'))

    if y['tblFinalizare']['CalePlanIncadrareDS']:
        pdf_list.insert(-1, y['tblFinalizare']['CalePlanIncadrareDS'].strip('"'))

    if y['tblFinalizare']['CalePlanSituatieDS']:
        pdf_list.insert(-1, y['tblFinalizare']['CalePlanSituatieDS'].strip('"'))

    if y['tblFinalizare']['CaleRaspunsUatDS']:
        pdf_list.insert(-1, y['tblFinalizare']['CaleRaspunsUatDS'].strip('"'))
    ### ----------------------------------------------------------------------------------------------------------------####

    if y['tblFinalizare']['CaleImputernicireDelgaz']:
        pdf_list_print.insert(2, y['tblFinalizare']['CaleImputernicireDelgaz'].strip('"'))

    if y['tblFinalizare']['CaleReferatDS']:
        pdf_list_print.insert(-1, y['tblFinalizare']['CaleReferatDS'].strip('"'))

    if y['tblFinalizare']['CaleRaportProiectant']:
        pdf_list_print.insert(-1, y['tblFinalizare']['CaleRaportProiectant'].strip('"'))

    x.merge_pdfs_print(pdf_list_print, path_document_final_print)






    ### ----------------------------------------------------------- creez EMAILUL  --------------------------------------------------####

    model_email = (r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\Executie\03.Pentru finalizare\01. Anunt ITL\Model email - provincie.docx")

    context_email = {
        'nume_client': y['client']['nume'],
        'email_uat': y['EmitentAC']['email'],
        'nr_ac': y['tblIncepereExecutie']['NumarAC'],
        'data_ac': x.get_date(y['tblIncepereExecutie']['DataAC']),
        'nume_lucrare': y['lucrare']['nume'],
        'nume_client': y['client']['nume'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
    }

    x.create_email(model_email, context_email, y['final_destination'])


    #### --------------------------------------------------------- facem curatenie -------------------------------------------###

    if os.path.exists(cerere_uat_pdf_path):
        os.remove(cerere_uat_pdf_path)

    if os.path.exists(declaratie_itl_pdf_path):
        os.remove(declaratie_itl_pdf_path)


    print("\n Anuntul UAT + ITL a fost creat \n")


def anunt_UAT_provincie(id_lucrare, path_final):
    director_final = '03. Anunt UAT - PROVINCIE - Finalizare'
    cale_stampila = "G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/DOCUMENTE/Stampila - RGT.png"
    y = x.get_data_finalizare(path_final, director_final, id_lucrare)
    valoare_reala = float(y['tblFinalizare']['ValoareReala'])   # valoarea fara proiectare

    model_cerere_uat = ('G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/Executie/03.Pentru finalizare/02. Anunt UAT/'f"Anunt UAT{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")
    
    context_cerere_uat = {
        'emitent_ac': y['EmitentAC']['denumire_institutie'].upper(),

        'nume_client': y['client']['nume'],

        'nume_firma_executie': y['firma_executie']['nume'],
        'nr_reg_com': y['firma_executie']['NrRegCom'],
        'cui_firma_executie': y['firma_executie']['CUI'],
        'localitate_firma_executie': y['firma_executie']['localitate'],
        'adresa_firma_executie': y['firma_executie']['adresa'],
        'judet_firma_executie': y['firma_executie']['judet'],
        'repr_firma_executie': y['firma_executie']['reprezentant'],
        'serie_ci': y['firma_executie']['seria_CI'],
        'nr_ci': y['firma_executie']['nr_CI'],
        'cnp_repr': y['firma_executie']['cnp_repr'],
        'telefon_contact': y['contact']['telefon'],
        'email_firma_executie': y['firma_executie']['email'],

        'nr_ac': y['tblIncepereExecutie']['NumarAC'],
        'data_ac': x.get_date(y['tblIncepereExecutie']['DataAC']),

        'data_finalizare': x.get_date(y['tblFinalizare']['DataFinalizare']),

        'nume_lucrare': y['lucrare']['nume'],

        'valoare_reala': f"{valoare_reala:.2f}",
        # Data
        'data': y['astazi'],
    }

    cerere_pdf_path = x.create_document(model_cerere_uat, context_cerere_uat, y['final_destination'], cale_stampila)
    path_document_final = os.path.join(path_final, director_final, f"Anunț UAT - Finalizare lucrări - {y['client']['nume']}.pdf")

    pdf_list = [
        cerere_pdf_path,
        y['firma_executie']['CaleCertificat'].strip('"'),
        y['firma_executie']['CaleCI'].strip('"'),
        y['tblIncepereExecutie']['CaleACScanat'].strip('"'),
        y['tblIncepereExecutie']['CalePlanIncadrareACScanat'].strip('"'),
        y['tblIncepereExecutie']['CalePlanSituatieACScanat'].strip('"'),
        y['tblIncepereExecutie']['CaleMemoriuTehnicACScanat'].strip('"'),
    ]

    if y['tblFinalizare']['CaleImputernicireDelgaz']:
        pdf_list.insert(2, y['tblFinalizare']['CaleImputernicireDelgaz'].strip('"'))

    if y['tblFinalizare']['CaleAdeverintaISC']:
        pdf_list.append(y['tblFinalizare']['CaleAdeverintaISC'].strip('"'))
    if y['tblFinalizare']['CaleITL']:
        pdf_list.append(y['tblFinalizare']['CaleITL'].strip('"'))
    if y['tblFinalizare']['CaleDovadaRegularizareTaxaAC']:
        pdf_list.append(y['tblFinalizare']['CaleDovadaRegularizareTaxaAC'].strip('"'))
    ### --------------------------- doar cand avem Dispozitie de santier --------------------------------------------- ###
    if y['tblFinalizare']['CaleDispozitieSantier']:
        pdf_list.insert(-1, y['tblFinalizare']['CaleDispozitieSantier'].strip('"'))
    if y['tblFinalizare']['CalePlanIncadrareDS']:
        pdf_list.insert(-1, y['tblFinalizare']['CalePlanIncadrareDS'].strip('"'))
    if y['tblFinalizare']['CalePlanSituatieDS']:
        pdf_list.insert(-1, y['tblFinalizare']['CalePlanSituatieDS'].strip('"'))
    if y['tblFinalizare']['CaleRaspunsUatDS']:
        pdf_list.insert(-1, y['tblFinalizare']['CaleRaspunsUatDS'].strip('"'))
    ### ----------------------------------------------------------------------------------------------------------------####
    if y['tblFinalizare']['CaleReferatDS']:
        pdf_list.append(y['tblFinalizare']['CaleReferatDS'].strip('"'))
    if y['tblFinalizare']['CaleRaportProiectant']:
        pdf_list.append(y['tblFinalizare']['CaleRaportProiectant'].strip('"'))

    x.merge_pdfs_print(pdf_list, path_document_final)

    if os.path.exists(cerere_pdf_path):
        os.remove(cerere_pdf_path)

        ### ----------------------------------------------------------- creez EMAILUL  --------------------------------------------------####

    model_email = (r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\Executie\03.Pentru finalizare\02. Anunt UAT\Model email - Anunt UAT.docx")

    context_email = {
        'nume_client': y['client']['nume'],
        'email_uat': y['EmitentAC']['email'],
        'nr_ac': y['tblIncepereExecutie']['NumarAC'],
        'data_ac': x.get_date(y['tblIncepereExecutie']['DataAC']),
        'nume_lucrare': y['lucrare']['nume'],
        'nume_client': y['client']['nume'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
    }

    x.create_email(model_email, context_email, y['final_destination'])

    print("\n Anuntul UAT PROVINCIE a fost creat \n")



def declaratie_ITL_provincie(id_lucrare, path_final):
    director_final = '03. Declaratie ITL - PROVINCIE'
    cale_stampila = "G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/DOCUMENTE/Stampila - RGT.png"
    y = x.get_data_finalizare(path_final, director_final, id_lucrare)
    valoare_reala = float(y['tblFinalizare']['ValoareReala'])   # valoarea fara proiectare

    if y['tblCU']['EmitentCU'] == 2:
        model_declaratie_itl = ('G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/Executie/03.Pentru finalizare/01. Anunt ITL/'f"ITL-064 - Tomesti{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")
    elif y['tblCU']['EmitentCU'] == 3:
        model_declaratie_itl = ('G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/Executie/03.Pentru finalizare/01. Anunt ITL/'f"ITL-064 - Miroslava{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")
    elif y['tblCU']['EmitentCU'] == 7:
        model_declaratie_itl = ('G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/Executie/03.Pentru finalizare/01. Anunt ITL/'f"ITL-064 - Barnova{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")
    elif y['tblCU']['EmitentCU'] == 15:
        model_declaratie_itl = ('G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/Executie/03.Pentru finalizare/01. Anunt ITL/'f"ITL-064 - Vama Suceava{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")
    elif y['tblCU']['EmitentCU'] == 14:
        model_declaratie_itl = ('G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/Executie/03.Pentru finalizare/01. Anunt ITL/'f"ITL-064 - Ciurea{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")
    else:
        print("NU am model de cerere ITL pentru aceastra localitate")

    expirare_executie = x.aduna_luni(y['tblIncepereExecutie']['DataIncepereExecutie'], int(y['tblIncepereExecutie']['ValabilitateExecutie']))
    valoare_ac = float(y['tblIncepereExecutie']['ValoareAC'])
    taxa_ac = x.custom_round(valoare_ac)
    taxa_reala = x.custom_round(valoare_reala * 0.01)
    diferenta_taxa = x.diferenta_taxa(taxa_ac, taxa_reala)
    taxa_reala = x.taxa_reala(valoare_reala, valoare_ac)

    context_declaratie_itl = {
        'nume_client': y['client']['nume'],
        'nume_firma_executie': y['firma_executie']['nume'],
        'nume_manager_proiect': y['manager_proiect']['nume'].upper(),
        'nr_ac': y['tblIncepereExecutie']['NumarAC'],
        'data_ac': x.get_date(y['tblIncepereExecutie']['DataAC']),
        'valoare_ac': f"{y['tblIncepereExecutie']['ValoareAC']:.2f}",
        'emitent_ac': y['EmitentAC']['denumire_institutie'],
        'expirare_executie': expirare_executie,

        'valoare_reala': f"{valoare_reala:.2f}",
        'taxa_reala': taxa_reala,
        'taxa_ac': f"{taxa_ac:.2f}", 
        'diferenta_taxa': diferenta_taxa,
        # Data
        'data': y['astazi'],
    }

    declaratie_itl_pdf_path = x.create_document(model_declaratie_itl, context_declaratie_itl, y['final_destination'], cale_stampila)
    path_document_final = os.path.join(path_final, director_final, f"Declaratie ITL - Finalizare lucrări - {y['client']['nume']}.pdf")


    pdf_list = [
        declaratie_itl_pdf_path,
        y['firma_executie']['CaleCertificat'].strip('"'),
        y['firma_executie']['CaleCI'].strip('"'),

        y['tblIncepereExecutie']['CaleACScanat'].strip('"'),
        y['tblIncepereExecutie']['CalePlanIncadrareACScanat'].strip('"'),
        y['tblIncepereExecutie']['CalePlanSituatieACScanat'].strip('"'),
        y['tblIncepereExecutie']['CaleMemoriuTehnicACScanat'].strip('"'),

        y['tblIncepereExecutie']['CaleContractRacordare'].strip('"'),
        y['tblIncepereExecutie']['CaleContractExecutie'].strip('"'),
        y['tblFinalizare']['CaleDevizFinal'].strip('"'),

        y['tblFinalizare']['CaleReferatDS'].strip('"'),
        y['tblFinalizare']['CaleRaportProiectant'].strip('"'),
        
    ]

    if y['tblFinalizare']['CaleImputernicireDelgaz']:
        pdf_list.insert(1, y['tblFinalizare']['CaleImputernicireDelgaz'].strip('"'))
    ### --------------------------- doar cand avem Dispozitie de santier --------------------------------------------- ###
    if y['tblFinalizare']['CaleDispozitieSantier']:
        pdf_list.insert(-1, y['tblFinalizare']['CaleDispozitieSantier'].strip('"'))
    if y['tblFinalizare']['CalePlanIncadrareDS']:
        pdf_list.insert(-1, y['tblFinalizare']['CalePlanIncadrareDS'].strip('"'))
    if y['tblFinalizare']['CalePlanSituatieDS']:
        pdf_list.insert(-1, y['tblFinalizare']['CalePlanSituatieDS'].strip('"'))
    if y['tblFinalizare']['CaleRaspunsUatDS']:
        pdf_list.insert(-1, y['tblFinalizare']['CaleRaspunsUatDS'].strip('"'))
    ### ----------------------------------------------------------------------------------------------------------------####

    if y['tblIncepereExecutie']['CaleDovadaPlataAC']:
        pdf_list.append(y['tblIncepereExecutie']['CaleDovadaPlataAC'].strip('"'))
    if y['tblIncepereExecutie']['CaleDovadaPlataISC']:
        pdf_list.append(y['tblIncepereExecutie']['CaleDovadaPlataISC'].strip('"'))

    if y['tblFinalizare']['CaleFacturiRGT']:
        pdf_list.append(y['tblFinalizare']['CaleFacturiRGT'].strip('"'))
    if y['tblFinalizare']['CaleDovadaPlataFacturi']:
        pdf_list.append(y['tblFinalizare']['CaleDovadaPlataFacturi'].strip('"'))


    x.merge_pdfs(pdf_list, path_document_final)

    if os.path.exists(declaratie_itl_pdf_path):
        os.remove(declaratie_itl_pdf_path)

    ### ----------------------------------------------------------- creez EMAILUL  --------------------------------------------------####

    model_email = (r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\Executie\03.Pentru finalizare\01. Anunt ITL\Model email - Declaratie ITL.docx")

    context_email = {
        'nume_client': y['client']['nume'],
        'email_uat': y['EmitentAC']['email'],
        'nr_ac': y['tblIncepereExecutie']['NumarAC'],
        'data_ac': x.get_date(y['tblIncepereExecutie']['DataAC']),
        'nume_lucrare': y['lucrare']['nume'],
        'nume_client': y['client']['nume'],
        'persoana_contact': y['contact']['nume'],
        'telefon_contact': y['contact']['telefon'],
    }

    x.create_email(model_email, context_email, y['final_destination'])



    print("\n Declaratia ITL PROVINCIE a fost creată \n")