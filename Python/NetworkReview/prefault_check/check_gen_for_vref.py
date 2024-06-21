def check_gen_for_vref(OlxObj,
                       expected_gen_vref=1.0):
    
    """
    Check for generator models with non-uniform internal
    voltage source magnitude

    Inputs:
        OlxObj:
            instance of the ASPEN OneLiner OlxObj API
            where the OLCase object is already loaded

        expected_gen_vref:
            float of expected generator internal voltage
            source magnitude

    Outputs:
        summary: A printable string containing basic summary information
                 about the checking results
        detail_dict: A dictionary containing the checking results details
    """

    detail_dict = {}
    detail_index = 0

    # vref range
    vref_range = sorted(list(set([eval(gen.PARAMSTR)['REFV'] for gen in OlxObj.OLCase.GEN])))
    vref_counts = {x:0 for x in vref_range}
    for gen in OlxObj.OLCase.GEN:
            gen_vref = eval(gen.PARAMSTR)['REFV']
            vref_counts[gen_vref] += 1

            if not gen_vref == expected_gen_vref:
                detail_dict[detail_index] = {'function':'check_gen_for_vref(expected_gen_vref=%s)'%(str(expected_gen_vref)),
                                            'object':gen.toString(),
                                            'area':gen.BUS.AREANO,
                                            'zone':gen.BUS.ZONENO,
                                            'condition':'gen_vref',
                                            'quantity':'gen.vref.pu',
                                            'expected':expected_gen_vref,
                                            'actual':gen_vref,
                                            'notes':'Inconsistent generator vref settings may cause unwanted prefault flows'}
                detail_index += 1
            
    summary = 'Found %s generator(s) with voltage reference different than %s pu'%(str(len(detail_dict)),expected_gen_vref)

    return summary,detail_dict

def main():
    pass


