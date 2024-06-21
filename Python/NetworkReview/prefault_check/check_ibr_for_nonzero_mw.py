def check_ibr_for_nonzero_mw(OlxObj):

    """
    Check IBR models for non-zero megawatt dispatch settings

    Inputs:
        OlxObj:
            instance of the ASPEN OneLiner OlxObj API
            where the OLCase object is already loaded

    Outputs:
        summary: A printable string containing basic summary information
                 about the checking results
        detail_dict: A dictionary containing the checking results details
    """
    
    detail_dict = {}
    detail_index = 0

    # check for load setting and total load
    load_ingored = bool(int(OlxObj.OLCase.getSystemParams()['bIgnoreLoad']))
    if load_ingored:
        load_setting_string = 'load is ignored'
    else:
        load_setting_string = 'load is not ignored'

    total_network_load_mw_abs = 0.0
    total_network_load_mvar_abs = 0.0
    for obj in OlxObj.OLCase.LOADUNIT:
        total_network_load_mw_abs += sum(obj.MW)
        total_network_load_mvar_abs += sum(obj.MVAR)
    total_network_load_mw_abs = round(total_network_load_mw_abs,2)
    total_network_load_mvar_abs = round(total_network_load_mvar_abs,2)
    total_network_load_mva_abs = (total_network_load_mw_abs**2 + total_network_load_mvar_abs**2)**(0.5)
    total_network_load_mva_abs = round(total_network_load_mva_abs,2)
    

    # cir
    for gen in OlxObj.OLCase.GENW4:
        if not gen.MW == 0:

            detail_dict[detail_index] = {'function':'check_ibr_for_nonzero_mw()',
                                           'object':gen.toString(),
                                           'area':gen.BUS.AREANO,
                                           'zone':gen.BUS.ZONENO,
                                           'condition':'non_zero_mw',
                                           'quantity':'gen.power.mw',
                                           'expected':0.0,
                                           'actual':gen.MW,
                                           'notes':'%s MVA of load in network and %s'%(total_network_load_mva_abs,load_setting_string)}
            detail_index += 1

    # type-3
    for gen in OlxObj.OLCase.GENW3:
        if not gen.MW == 0:

            detail_dict[detail_index] = {'function':'check_ibr_for_nonzero_mw()',
                                           'object':gen.toString(),
                                           'area':gen.BUS.AREANO,
                                           'zone':gen.BUS.ZONENO,
                                           'condition':'non_zero_mw',
                                           'quantity':'gen.power.mw',
                                           'expected':0.0,
                                           'actual':gen.MW,
                                           'notes':'%s MVA of load in network and %s'%(total_network_load_mva_abs,load_setting_string)}
            detail_index += 1
    
    summary = 'Found %s CIR and/or Type-3 model(s) with non-zero MW dispatch'%(str(len(detail_dict)))

    return summary,detail_dict

def main():
    pass

