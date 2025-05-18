import os
import functii as x
from dateutil.relativedelta import relativedelta




id_lucrare = 63
path_final = r"G:\Shared drives\Root\11. DATABASE\Pentru CRISTI\Test"


def declaratie_ITL(id_lucrare, path_final):
    director_final = '01. Declaratie ITL'
    cale_stampila = "G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/DOCUMENTE/Stampila - RGT.png"
    y = x.get_data_finalizare(path_final, director_final, id_lucrare)


    model_cerere = ('G:/Shared drives/Root/11. DATABASE/01. Automatizari avize/Executie/03.Pentru finalizare/01. Anunt ITL/'f"ITL-064 - Iasi{' - DELGAZ' if y['lucrare']['IDClient'] == 1 else ''}.docx")
 
    expirare_executie = x.aduna_luni(y['tblIncepereExecutie']['DataIncepereExecutie'], int(y['tblIncepereExecutie']['ValabilitateExecutie']))
    valoare_reala = float(y['tblFinalizare']['ValoareReala'])  # fara proiectare    
    taxa_ac = x.custom_round(float(y['tblIncepereExecutie']['ValoareAC']) * 0.01)
    taxa_reala = x.custom_round(valoare_reala * 0.01)
    diferenta_taxa = x.diferenta_taxa(taxa_ac, taxa_reala)

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
        'taxa_reala': f"{taxa_reala:.2f}",
        'taxa_ac': f"{taxa_ac:.2f}", 
        'diferenta_taxa': diferenta_taxa,
        # Data
        'data': y['astazi'],
    }

    cerere_pdf_path = x.create_document(model_cerere, context_cerere, y['final_destination'], cale_stampila)

    path_document_final = os.path.join(path_final, director_final, f"Documentatie ITL - {y['client']['nume']}.pdf")

    print(47)

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

    print(68)

    if y['tblFinalizare']['CaleImputernicireDelgaz']:
        pdf_list.insert(1, y['tblFinalizare']['CaleImputernicireDelgaz'].strip('"'))
    print(72)
    ### --------------------------- doar cand avem Dispozitie de santier --------------------------------------------- ###
    if y['tblFinalizare']['CaleDispozitieSantier']:
        pdf_list.insert(-1, y['tblFinalizare']['CaleDispozitieSantier'].strip('"'))

    print(77)
    if y['tblFinalizare']['CalePlanIncadrareDS']:
        pdf_list.insert(-1, y['tblFinalizare']['CalePlanIncadrareDS'].strip('"'))

    print(81)
    if y['tblFinalizare']['CalePlanSituatieDS']:
        pdf_list.insert(-1, y['tblFinalizare']['CalePlanSituatieDS'].strip('"'))

    print(85)
    if y['tblFinalizare']['CaleRaspunsUatDS']:
        pdf_list.insert(-1, y['tblFinalizare']['CaleRaspunsUatDS'].strip('"'))

    print(89)
    ### ----------------------------------------------------------------------------------------------------------------####

    if y['tblIncepereExecutie']['CaleDovadaPlataAC']:
        pdf_list.append(y['tblIncepereExecutie']['CaleDovadaPlataAC'].strip('"'))
    print(94)

    if y['tblIncepereExecutie']['CaleDovadaPlataISC']:
        pdf_list.append(y['tblIncepereExecutie']['CaleDovadaPlataISC'].strip('"'))

    print(99)

    if y['tblFinalizare']['CaleFacturiRGT']:
        pdf_list.append(y['tblFinalizare']['CaleFacturiRGT'].strip('"'))
    
    print(104)
    if y['tblFinalizare']['CaleDovadaPlataFacturi']:
        pdf_list.append(y['tblFinalizare']['CaleDovadaPlataFacturi'].strip('"'))

    print(108)
    x.merge_pdfs_signed(pdf_list, path_document_final)

    print(111)
    if os.path.exists(cerere_pdf_path):
        os.remove(cerere_pdf_path)


    # ### ------------------------------------------------------------- Creez EMAILUL ------------------------------------------------ ###


    # model_email = (r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\Executie\03.Pentru finalizare\01. Anunt ITL\Model email.docx")

    # context_email = {
    #     'nume_client': y['client']['nume'],
    #     'nr_ac': y['tblIncepereExecutie']['NumarAC'],
    #     'data_ac': x.get_date(y['tblIncepereExecutie']['DataAC']),
    #     'nume_lucrare': y['lucrare']['nume'],

    #     'persoana_contact': y['contact']['nume'],
    #     'telefon_contact': y['contact']['telefon'],
    # }

    # x.create_email(model_email, context_email, y['final_destination'])


    print("\n Declaratia ITL a fost creată \n")



declaratie_ITL(id_lucrare, path_final)