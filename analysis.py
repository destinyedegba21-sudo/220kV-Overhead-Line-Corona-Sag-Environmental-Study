"""Independent study calculations; never writes to the source workbook."""
from pathlib import Path
import argparse
import hashlib
import math
import textwrap

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from openpyxl import load_workbook


# Absolute tolerances in the reported units; no relative tolerance is used.
TOLERANCES = {'m': 1e-8, 'mm': 1e-5, 'kV': 1e-7, 'kV LL': 1e-7,
              'kV/cm': 1e-8, '%': 1e-7, 'A': 1e-7, 'N': 1e-5,
              'kN': 1e-8, 'N/m': 1e-8, 'Pa': 1e-7, 'deg': 1e-7,
              'mH/km': 1e-9, 'H': 1e-10, 'ohm/phase': 1e-7,
              'nF/km': 1e-8, 'uF': 1e-9, 'kW': 1e-6,
              'kW/km/phase': 1e-8, 'degC': 1e-8, '1': 1e-9}


def sag(weight, span, tension):
    """Parabolic sag with distributed load N/m, span m and horizontal tension N."""
    return weight * span**2 / (8 * tension)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('workbook', nargs='?', type=Path,
                        default=Path(__file__).with_name('Project_4_220kV_Corona_Sag_Environmental_Study.xlsx'))
    args = parser.parse_args()
    source = args.workbook.resolve()
    before = hashlib.sha256(source.read_bytes()).hexdigest()
    # Read-only mode bypasses an incompatible chart definition in this workbook.
    wb = load_workbook(source, read_only=True, data_only=True)
    formulas = load_workbook(source, read_only=True, data_only=False)
    rows = []

    def value(sheet, cell):
        return wb[sheet][cell].value

    def raw(sheet, cell):
        result = formulas[sheet][cell].value
        if not isinstance(result, (int, float)):
            raise ValueError(f'Expected a literal study input at {sheet}!{cell}: {result}')
        return float(result)

    def record(parameter, scenario, result, sheet, cell, unit):
        excel = value(sheet, cell)
        available = isinstance(excel, (int, float)) and math.isfinite(excel)
        difference = float(result) - excel if available else np.nan
        status = ('MATCH' if abs(difference) <= TOLERANCES[unit] else 'DIFFERENCE') if available else 'NO EXCEL CACHE'
        rows.append([parameter, f'{scenario} [{sheet}!{cell}]', float(result),
                     excel if available else np.nan, difference, unit, status])

    inp = lambda row: raw('01_Inputs', f'B{row}')
    voltage, frequency, length = inp(5), inp(6), inp(7)
    # GMD spacings, pressure and 273 convention are literals in input formulas.
    # Refuse a changed formula rather than silently using stale assumptions.
    expected = {'B8': '=(6*6*12)^(1/3)', 'B23': '=(3.92*76)/(273+25)',
                'B24': '=(3.92*76)/(273+40)'}
    for cell, formula in expected.items():
        if formulas['01_Inputs'][cell].value != formula:
            raise ValueError(f'Review changed formula 01_Inputs!{cell} before running.')
    gmd = (6 * 6 * 12)**(1/3)
    radius, mass, separation = inp(9), inp(10), inp(11)
    tension, span, height = inp(12), inp(13), inp(14)
    reference_t, alpha = inp(16), inp(18)
    rho, drag = inp(20), inp(21)
    delta = lambda t: 3.92 * 76 / (273 + t)
    phase = voltage / math.sqrt(3)
    critical = lambda r_cm, surface, t=25: inp(27)*surface*delta(t)*r_cm*math.log(gmd*100/r_cm)
    dry, wet = [critical(radius*100, inp(i)) for i in (25, 26)]
    loss = lambda v, threshold: inp(31)/delta(25)*(frequency+inp(30))*math.sqrt(radius/gmd)*np.maximum(np.asarray(v)/math.sqrt(3)-threshold, 0)**2*1e-5
    sheet = '02_Corona_Analysis'
    record('Operating phase voltage', 'Baseline', phase, sheet, 'B5', 'kV')
    for col, name, threshold in [('B', 'Dry', dry), ('C', 'Wet', wet)]:
        for row, label, result, unit in [
            (6, 'Air-density factor', delta(25), '1'),
            (8, 'Critical gradient', inp(27)*inp(25 if col=='B' else 26)*delta(25), 'kV/cm'),
            (9, 'Disruptive critical voltage, phase', threshold, 'kV'),
            (10, 'Disruptive critical voltage, LL', threshold*math.sqrt(3), 'kV LL'),
            (11, 'Inception margin, phase', threshold-phase, 'kV'),
            (12, 'Inception margin, LL', threshold*math.sqrt(3)-voltage, 'kV LL'),
            (13, 'Corona loss', loss(voltage, threshold), 'kW/km/phase'),
            (14, 'Total corona loss', 3*length*loss(voltage, threshold), 'kW')]:
            record(label, name, result, sheet, f'{col}{row}', unit)
    voltages = [raw(sheet, f'A{r}') for r in range(18, 23)]
    for r, v in zip(range(18, 23), voltages):
        for c, result, unit in [('B', v/math.sqrt(3), 'kV'), ('C', loss(v,dry), 'kW/km/phase'), ('D',loss(v,wet),'kW/km/phase'), ('E',3*length*loss(v,dry),'kW'), ('F',3*length*loss(v,wet),'kW')]:
            record('Voltage sensitivity '+c, f'{v:g} kV LL', result, sheet, f'{c}{r}', unit)
    for r,t in [(26,25),(27,40)]:
        visual = critical(radius*100,inp(25),t)*(1+0.3/math.sqrt(delta(t)*radius*100))
        record('Visual critical voltage, phase',f'{t} degC',visual,sheet,f'B{r}','kV')
        record('Visual critical voltage, LL',f'{t} degC',visual*math.sqrt(3),sheet,f'C{r}','kV LL')
    for r,t in [(32,25),(33,40)]:
        threshold = critical(radius*100,inp(25),t)*math.sqrt(3)
        for c, result, unit in [('B',threshold,'kV LL'),('C',threshold-voltage,'kV LL'),('D',(threshold-voltage)/threshold*100,'%')]:
            record('Temperature inception margin '+c,f'{t} degC',result,sheet,f'{c}{r}',unit)
    radii = [raw(sheet,f'A{r}') for r in range(37,40)]
    for r,rc in zip(range(37,40),radii):
        threshold=critical(rc,inp(25))
        for c,result,unit in [('B',threshold,'kV'),('C',threshold*math.sqrt(3),'kV LL'),('D',(threshold-dry)/dry*100,'%')]:
            record('Radius sensitivity '+c,f'{rc:g} cm',result,sheet,f'{c}{r}',unit)

    # Bundling: different equivalent radii for inductance and capacitance.
    sheet='03_Bundling'
    gmr=radius*inp(29)
    bundle_gmr=math.sqrt(gmr*separation)
    bundle_radius=math.sqrt(radius*separation)
    enhancement=1+radius/separation
    electrical=[]
    for col,name,rg,rc,n,enh in [('B','Single',gmr,radius,1,1),('C','Twin bundle',bundle_gmr,bundle_radius,2,enhancement)]:
        field=phase/(n*radius*math.log(gmd/rc))/100
        onset=inp(27)*inp(25)*delta(25)*n*radius*100*math.log(gmd/rc)/enh
        inductance=2e-7*math.log(gmd/rg)*1e6  # H/m -> mH/km
        total_l=inductance*length/1000
        cap=2*math.pi*inp(28)/math.log(gmd/rc)*1e12  # F/m -> nF/km
        total_c=cap*length/1000  # nF -> uF
        current=2*math.pi*frequency*total_c*phase/1000
        electrical.append((inductance,cap,field*enh,onset))
        for r,label,result,unit in [(7,'Capacitance equivalent radius',rc,'m'),(8,'Inductance GMR',rg,'m'),(9,'Capacitance radius ratio',rc/radius,'1'),(10,'GMR ratio',rg/gmr,'1'),(11,'Critical electric field',inp(27)*inp(25)*delta(25),'kV/cm'),(12,'Surface electric field',field,'kV/cm'),(13,'Field enhancement',enh,'1'),(14,'Maximum electric field',field*enh,'kV/cm'),(15,'Corona inception phase voltage',onset,'kV'),(16,'Corona inception LL voltage',onset*math.sqrt(3),'kV LL'),(17,'Inductance per km',inductance,'mH/km'),(18,'Total inductance',total_l,'H'),(19,'Inductive reactance',2*math.pi*frequency*total_l,'ohm/phase'),(20,'Capacitance per km',cap,'nF/km'),(21,'Total capacitance',total_c,'uF'),(22,'Charging current',current,'A')]:
            record(label,name,result,sheet,f'{col}{r}',unit)
    single,twin=electrical
    changes=[(single[0]-twin[0])/single[0]*100,(twin[1]-single[1])/single[1]*100,(single[2]-twin[2])/single[2]*100,(twin[3]-single[3])/single[3]*100]
    labels=['Inductance reduction','Capacitance increase','Electric-field reduction','Corona inception improvement']
    for r,label,result in zip(range(27,31),labels,changes):
        record(label,'Twin versus single',result,sheet,f'B{r}','%')

    # Mechanical baseline and workbook sensitivity grids.
    weight=mass*inp(15)
    baseline=sag(weight,span,tension)
    reaction=weight*span/2
    sheet='04_Sag_Tension'
    for r,label,result,unit in [(5,'Conductor weight',weight,'N/m'),(6,'Baseline conductor sag',baseline,'m'),(7,'Baseline ground clearance',height-baseline,'m'),(8,'Sag/span',baseline/span*100,'%'),(9,'Vertical support reaction',reaction,'N'),(10,'Resultant support tension',math.hypot(tension,reaction),'N')]:
        record(label,'Baseline',result,sheet,f'B{r}',unit)
    spans=[raw(sheet,f'A{r}') for r in range(14,19)]
    tensions=[raw(sheet,f'A{r}') for r in range(23,28)]
    for r,L in zip(range(14,19),spans):
        d=sag(weight,L,tension)
        for c,result,unit in [('B',d,'m'),('C',height-d,'m'),('D',100*d/L,'%'),('E',weight*L/2000,'kN')]:
            record('Span sensitivity '+c,f'{L:g} m',result,sheet,f'{c}{r}',unit)
    for r,H in zip(range(23,28),tensions):
        d=sag(weight,span,H*1000)
        for c,result,unit in [('B',d,'m'),('C',height-d,'m'),('D',math.hypot(H*1000,reaction)/1000,'kN')]:
            record('Tension sensitivity '+c,f'{H:g} kN',result,sheet,f'{c}{r}',unit)
    for r in range(32,37):
        extra=raw(sheet,f'A{r}'); w=weight*(1+extra/100); d=sag(w,span,tension)
        for c,result,unit in [('B',w,'N/m'),('C',d,'m'),('D',height-d,'m'),('E',math.hypot(tension,w*span/2)/1000,'kN')]:
            record('Additional loading '+c,f'{extra:g} %',result,sheet,f'{c}{r}',unit)

    sheet='05_Temperature'
    suspended=span+8*baseline**2/(3*span)
    # The standalone block uses literal =25 and =B6+50, and a local alpha.
    if formulas[sheet]['B6'].value != '=25' or formulas[sheet]['B7'].value != '=B6+50':
        raise ValueError('Review changed standalone temperature formulas.')
    local_alpha=raw(sheet,'B9'); dt=50
    expansion=local_alpha*suspended*dt
    hot=math.sqrt(3*span*(suspended+expansion-span)/8)
    for r,label,result,unit in [(8,'Temperature increase',dt,'degC'),(10,'Span thermal expansion',local_alpha*span*dt,'m'),(11,'Span thermal expansion',local_alpha*span*dt*1000,'mm'),(12,'Expanded reference length',span*(1+local_alpha*dt),'m'),(19,'Initial suspended length',suspended,'m'),(20,'Suspended thermal expansion',expansion*1000,'mm'),(21,'Expanded suspended length',suspended+expansion,'m'),(22,'Hot conductor sag',hot,'m'),(23,'Hot clearance',height-hot,'m'),(24,'Sag increase',(hot-baseline)/baseline*100,'%')]:
        record(label,'Standalone 75 degC',result,sheet,f'B{r}',unit)
    temperatures=[]; temperature_sags=[]; clearances=[]
    for r in range(29,34):
        t=raw(sheet,f'A{r}'); temperatures.append(t)
        increase=t-reference_t
        expansion=alpha*suspended*increase
        expanded=suspended+expansion
        d=math.sqrt(3*span*(expanded-span)/8)
        temperature_sags.append(d); clearances.append(height-d)
        for c,result,unit in [('B',increase,'degC'),('C',expansion,'m'),('D',expanded,'m'),('E',d,'m'),('F',height-d,'m'),('G',(d-baseline)/baseline*100,'%')]:
            record('Temperature sensitivity '+c,f'{t:g} degC',result,sheet,f'{c}{r}',unit)

    sheet='06_Wind_Loading'
    speeds=[]; pressures=[]
    for r in range(5,11):
        speed=raw(sheet,f'A{r}'); pressure=0.5*rho*speed**2
        wind=pressure*drag*2*radius; displacement=sag(wind,span,tension)
        speeds.append(speed); pressures.append(pressure)
        for c,result,unit in [('B',pressure,'Pa'),('C',wind,'N/m'),('D',math.hypot(weight,wind),'N/m'),('E',math.degrees(math.atan(wind/weight)),'deg'),('F',displacement,'m'),('G',math.hypot(baseline,displacement),'m')]:
            record('Wind sensitivity '+c,f'{speed:g} m/s',result,sheet,f'{c}{r}',unit)

    sheet='07_Environmental'; scenario_names=[]; environmental_clearances=[]
    hot_tension=weight*span**2/(8*hot)
    for r in range(6,13):
        name=value(sheet,f'A{r}'); t=raw(sheet,f'B{r}'); speed=raw(sheet,f'C{r}'); extra=raw(sheet,f'D{r}')
        scenario_names.append(name)
        # Exact selected tension policy: hot rows share G7, others nominal H.
        H=hot_tension if r in (7,11,12) else tension
        w=weight*(1+extra/100); wind=0.5*rho*speed**2*drag*2*radius
        vertical=sag(w,span,H); horizontal=sag(wind,span,H)
        resultant=math.hypot(vertical,horizontal)
        support=math.sqrt(H**2+(w*span/2)**2+(wind*span/2)**2)/1000
        clearance=height-vertical
        environmental_clearances.append(clearance)
        for c,result,unit in [('E',w,'N/m'),('F',wind,'N/m'),('G',H,'N'),('H',vertical,'m')]:
            record('Environmental loading '+c,name,result,sheet,f'{c}{r}',unit)
        out=r+11
        for c,label,result,unit in [('B','Horizontal displacement',horizontal,'m'),('C','Resultant displacement',resultant,'m'),('D','Ground clearance',clearance,'m'),('E','Resultant support tension',support,'kN'),('F','Clearance reduction',height-baseline-clearance,'m')]:
            record(label,name,result,sheet,f'{c}{out}',unit)

    outdir=Path(__file__).resolve().parent/'outputs'
    figdir=outdir/'figures'; figdir.mkdir(parents=True,exist_ok=True)
    plt.rcParams.update({'font.size':10,'axes.spines.top':False,'axes.spines.right':False,'figure.figsize':(8,4.8)})

    def graph(filename,title,xlabel,ylabel,x,series,bar=False):
        fig,ax=plt.subplots(figsize=(9,5.7) if bar else (8,4.8))
        if bar:
            ax.barh(np.arange(len(x)),series[0][1],color='#28739b')
            ax.set_yticks(np.arange(len(x)),[textwrap.fill(str(s),27) for s in x])
            ax.invert_yaxis()
        else:
            for label,y in series:
                ax.plot(x,y,marker='o',label=label)
            if len(series)>1: ax.legend(frameon=False)
        ax.set(title=title,xlabel=xlabel,ylabel=ylabel)
        ax.grid(axis='x' if bar else 'both',alpha=0.2)
        ax.set_axisbelow(True)
        fig.tight_layout(rect=(0,0,1,1))
        fig.savefig(figdir/filename,dpi=300)
        plt.close(fig)

    graph('01_corona_loss.png','Corona power loss versus operating voltage','Operating line-to-line voltage (kV)','Corona loss (kW/km/phase)',voltages,[('Dry',loss(voltages,dry)),('Wet',loss(voltages,wet))])
    graph('02_corona_radius.png','Conductor radius and disruptive critical voltage','Conductor radius (cm)','Disruptive critical voltage (kV LL)',radii,[('Dry at 25°C',[critical(r,inp(25))*math.sqrt(3) for r in radii])])
    graph('03_bundling_comparison.png','Twin-bundle performance versus single conductor','Change (%)','',labels,[('',changes)],bar=True)
    graph('04_sag_span.png','Conductor sag versus span length','Span length (m)','Conductor sag (m)',spans,[('',[sag(weight,L,tension) for L in spans])])
    graph('05_sag_tension.png','Conductor sag versus horizontal tension','Horizontal tension (kN)','Conductor sag (m)',tensions,[('',[sag(weight,span,H*1000) for H in tensions])])
    graph('06_temperature_sag.png','Conductor sag versus temperature','Conductor temperature (°C)','Conductor sag (m)',temperatures,[('',temperature_sags)])
    graph('07_temperature_clearance.png','Ground clearance versus temperature','Conductor temperature (°C)','Ground clearance (m)',temperatures,[('',clearances)])
    graph('08_wind_pressure.png','Wind pressure versus wind speed','Wind speed (m/s)','Wind pressure (Pa)',speeds,[('',pressures)])
    graph('09_environmental_clearance.png','Ground clearance by environmental scenario','Ground clearance (m)','',scenario_names,[('',environmental_clearances)],bar=True)
    df=pd.DataFrame(rows,columns=['Engineering Parameter','Scenario','Python Result','Excel Result','Difference','Unit','Verification Status'])
    # Preserve precision in the CSV so small differences remain inspectable.
    df.to_csv(outdir/'results_summary.csv',index=False,float_format='%.15g')
    wb.close(); formulas.close()
    if hashlib.sha256(source.read_bytes()).hexdigest()!=before:
        raise RuntimeError('Source workbook checksum changed.')
    assert len(list(figdir.glob('*.png')))==9, 'Expected exactly nine PNG files.'
    print(df['Verification Status'].value_counts().to_string())
    print('Source SHA-256 unchanged:',before)
    print('Figures: 9; CSV:',outdir/'results_summary.csv')


if __name__=='__main__':
    main()
