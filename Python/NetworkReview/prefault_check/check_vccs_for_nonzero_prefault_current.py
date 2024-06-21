def check_vccs_for_nonzero_prefault_current(OlxObj,
                                            voltage_checkpoints_pu=(0.95,1.0,1.05)):
    
    """
    Check for VCCS that have non-zero currents defined in the
    normal voltage range portion of the V-I table.  The reported
    value is the maximum amount that could occur within the
    input voltage range, not the actual amount for a given
    prefault condition.

    Inputs:
        OlxObj:
            instance of the ASPEN OneLiner OlxObj API
            where the OLCase object is already loaded

        voltage_checkpoints_pu:
            tuple of per unit voltages that will be
            checked for non-zero current

    Outputs:
        summary: A printable string containing basic summary information
                 about the checking results
        detail_dict: A dictionary containing the checking results details
    """

    summary_dict = {}

    detail_dict = {}
    detail_index = 0

    # function to check vccs output magnitude
    def get_vccs_output(v,vccs_v,vccs_i,vccs_a):
        # init output
        vccs_i_out = 0.0

        # strip away all-zero rows
        v_i_dict = {vv:ii for vv,ii,aa in zip(vccs_v,vccs_i,vccs_a) if not all([vv == 0.0,
                                                                                ii == 0.0,
                                                                                aa == 0.0])}
        #print(v_i_dict)
        vccs_v_max = max(v_i_dict.keys())
        vccs_v_min = min(v_i_dict.keys())
        
        if v > vccs_v_max:
            if v_i_dict[vccs_v_max] == 0.0:
                vccs_i_out = 0.0
            else:
                vccs_i_out = v_i_dict[vccs_v_max] # i at max v
        elif v < vccs_v_min:
            if v_i_dict[vccs_v_min] == 0.0:
                vccs_i_out = 0.0
            else:
                vccs_i_out = v_i_dict[vccs_v_min] # i at min v
        else: # v in inside table domain
            # v directly specified in table
            if v in v_i_dict:
                vccs_i_out = v_i_dict[v]
            else:
                # find nearest values above and below
                for vv,ii in sorted(v_i_dict.items(),reverse=False):
                    if vv >= v:
                        v_up = vv
                        i_up = ii
                        break
                for vv,ii in sorted(v_i_dict.items(),reverse=True):
                    if vv <= v:
                        v_down = vv
                        i_down = ii
                        break
                # linear interpolation
                slope = (i_up - i_down) / (v_up - v_down)
                vccs_i_out = i_up - (v_up - v) * slope
        return vccs_i_out

    
    for gen in OlxObj.OLCase.CCGEN:
        i_max_voltage_checkpoints_pu = max([abs(get_vccs_output(vv,gen.V,gen.I,gen.A)) for vv in voltage_checkpoints_pu])
        if i_max_voltage_checkpoints_pu > 0.0:

            detail_dict[detail_index] = {'function':'check_vccs_for_nonzero_prefault_current(voltage_checkpoints_pu=%s)'%(str(voltage_checkpoints_pu)),
                                           'object':gen.toString(),
                                           'area':gen.BUS.AREANO,
                                           'zone':gen.BUS.ZONENO,
                                           'condition':'vccs_prefault_current',
                                           'quantity':'gen.current.amps',
                                           'expected':0.0,
                                           'actual':i_max_voltage_checkpoints_pu,
                                           'notes':'This is max possible flow within nominal range, not actual flow'}
            detail_index += 1
            
    summary = 'Found %s VCCS model(s) with non-zero current at voltage points %s pu'%(str(len(detail_dict)),str(voltage_checkpoints_pu))

    return summary,detail_dict

def main():
    pass

