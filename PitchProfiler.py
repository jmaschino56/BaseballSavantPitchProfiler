import pandas as pd
from pybaseball import statcast_pitcher
from pybaseball import playerid_lookup
import matplotlib.pylab as plt
import matplotlib.gridspec as gridspec
import math as m
from pandas.compat import BytesIO
from docx import Document
from docx.shared import Inches
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import numpy as np
import os

plt.style.use('seaborn-paper')

'''
# used for debugging
pd.set_option('display.max_columns', None)
pd.set_option('display.expand_frame_repr', False)
pd.set_option('max_colwidth', -1)
'''


def getNumber(last, first):
    playerTable = playerid_lookup(last, first)
    playerTable = playerTable.loc[playerTable['mlb_played_last'].isin([2019])]
    playerTable.index = range(len(playerTable['mlb_played_last']))
    number = playerTable['key_mlbam']
    number = number[0]
    return number


def dataGrab(number, start, end):
    data = statcast_pitcher(start_dt=start, end_dt=end,
                            player_id=number)
    data = data[['pitch_type', 'release_speed', 'release_pos_x', 'release_pos_z',
                 'pfx_x', 'pfx_z', 'release_spin_rate', 'plate_x', 'plate_z',
                 'estimated_woba_using_speedangle', 'woba_value', 'description',
                 'launch_speed_angle', 'launch_angle', 'launch_speed', 'bb_type',
                 'effective_speed']]
    data.index = range(len(data['pitch_type']))
    return data


def getPitchTypes(data):
    pitchcounts = data['pitch_type'].value_counts(dropna=True)
    pitch_types = []
    for i in range(len(pitchcounts)):
        is_pitch = pitchcounts.index[i]
        pitch_types.append(is_pitch)
    return pitch_types


def colorPicker(pitch_type):
    color = ''
    if(pitch_type == 'FF'):
        color = '#8C1C13'
    elif(pitch_type == 'FT'):
        color = '#CC5803'
    elif(pitch_type == 'SI'):
        color = '#FE7F2D'
    elif(pitch_type == 'FC'):
        color = '#E08E45'
    elif(pitch_type == 'FS'):
        color = '#F3CA4C'
    elif(pitch_type == 'SL'):
        color = '#274060'
    elif(pitch_type == 'CU'):
        color = '#4EA5D9'
    elif(pitch_type == 'KC'):
        color = '#5BC0EE'
    elif(pitch_type == 'CH'):
        color = '#1446A0'
    elif(pitch_type == 'KN'):
        color = '#712F79'
    elif(pitch_type == 'FO'):
        color = '#03B5AA'
    elif(pitch_type == 'EP'):
        color = '#DBFE97'
    elif(pitch_type == 'SC'):
        color = '#5B9279'
    else:
        color = 'black'
    return color


def plotData(data):
    pitch_types = getPitchTypes(data)
    gs = gridspec.GridSpec(2, 2)
    fig = plt.figure(figsize=(6, 2.5))
    ax0 = plt.subplot(gs[:, 0])  # release
    ax2 = plt.subplot(gs[:, 1])  # break
    for i in range(len(pitch_types)):
        is_pitch = data['pitch_type'] == pitch_types[i]
        selected_data = data[is_pitch]
        label = pitch_types[i]
        color = colorPicker(label)
        if(label == 'PO' or label == 'IB' or label == 'AB' or label == 'UN' or label == 'EP'):
            continue
        else:
            ax0.scatter(selected_data['release_pos_x'], selected_data['release_pos_z'],
                        label=label, s=20, alpha=0.5, c=color)
            ax2.scatter(12*selected_data['pfx_x'], 12*selected_data['pfx_z'],
                        label=label, s=20, alpha=0.5, c=color)
    ax0.set_xlim(data['release_pos_x'].mean()-1.5, data['release_pos_x'].mean()+1.5)
    ax0.set_ylim(data['release_pos_z'].mean()-1.5, data['release_pos_z'].mean()+1.5)
    ax0.set_title('Release Position')
    ax0.set_xlabel('Horizontal Release Point')
    ax0.set_ylabel('Vertical Release Point')
    ax2.set_xlim(-30, 30)
    ax2.set_ylim(-30, 30)
    ax2.axhline(y=0, color='k')
    ax2.axvline(x=0, color='k')
    ax2.set_title('Pitch Movement')
    ax2.set_xlabel('Horizontal Movement')
    ax2.set_ylabel('Vertical Movement')
    ax0.legend(prop={'size': 7})
    ax2.legend(prop={'size': 7})
    fig.tight_layout()
    fig.subplots_adjust(top=.92, wspace=0.35, bottom=.18, left=.11, right=.98)
    memfile = BytesIO()
    plt.savefig(memfile)
    # plt.show() for debug
    return memfile


def plotLocationData(data):
    pitch_types = getPitchTypes(data)
    gs = gridspec.GridSpec(1, 7)
    fig = plt.figure(figsize=(6, 1.5))
    ax0 = plt.subplot(gs[:, 0])  # FF
    ax1 = plt.subplot(gs[:, 1])  # FT/SI
    ax2 = plt.subplot(gs[:, 2])  # FC
    ax3 = plt.subplot(gs[:, 3])  # SL
    ax4 = plt.subplot(gs[:, 4])  # CU/KC
    ax5 = plt.subplot(gs[:, 5])  # CH
    ax6 = plt.subplot(gs[:, 6])  # FS

    for i in range(len(pitch_types)):
        is_pitch = data['pitch_type'] == pitch_types[i]
        selected_data = data[is_pitch]
        label = pitch_types[i]
        # color = colorPicker(label)
        sz_x = [.79, .79, -.79, -.79, .79]
        sz_z = [3.5, 1.5, 1.5, 3.5, 3.5]
        xedges, zedges = np.linspace(-2, 2, 20), np.linspace(-0.5, 4.5, 20)
        x = selected_data['plate_x']
        z = selected_data['plate_z']
        hist, xedges, yedges = np.histogram2d(x, z, (xedges, zedges))
        xidx = np.clip(np.digitize(x, xedges), 0, hist.shape[0]-1)
        zidx = np.clip(np.digitize(z, zedges), 0, hist.shape[1]-1)
        c = hist[xidx, zidx]
        if(label == 'FF'):
            ax0.scatter(x, z, c=c, s=1, cmap='YlOrRd')
            ax0.plot(sz_x, sz_z, color='black', lw=0.5)
        elif(label == 'FT' or label == 'SI'):
            ax1.scatter(x, z, c=c, s=1, cmap='YlOrRd')
            ax1.plot(sz_x, sz_z, color='black', lw=0.5)
        elif(label == 'FC'):
            ax2.scatter(x, z, c=c, s=1, cmap='YlOrRd')
            ax2.plot(sz_x, sz_z, color='black', lw=0.5)
        elif(label == 'SL'):
            ax3.scatter(x, z, c=c, s=1, cmap='YlOrRd')
            ax3.plot(sz_x, sz_z, color='black', lw=0.5)
        elif(label == 'CU' or label == 'KC'):
            ax4.scatter(x, z, c=c, s=1, cmap='YlOrRd')
            ax4.plot(sz_x, sz_z, color='black', lw=0.5)
        elif(label == 'CH'):
            ax5.scatter(x, z, c=c, s=1, cmap='YlOrRd')
            ax5.plot(sz_x, sz_z, color='black', lw=0.5)
        elif(label == 'FS'):
            ax6.scatter(x, z, c=c, s=1, cmap='YlOrRd')
            ax6.plot(sz_x, sz_z, color='black', lw=0.5)
    ax0.set_xlim(-2, 2)
    ax0.set_ylim(-0.5, 4.5)
    ax0.axis('off')
    ax0.set_title('FF', fontsize=10)
    ax1.set_xlim(-2, 2)
    ax1.set_ylim(-0.5, 4.5)
    ax1.axis('off')
    ax1.set_title('FT/SI', fontsize=10)
    ax2.set_xlim(-2, 2)
    ax2.set_ylim(-0.5, 4.5)
    ax2.axis('off')
    ax2.set_title('FC', fontsize=10)
    ax3.set_xlim(-2, 2)
    ax3.set_ylim(-0.5, 4.5)
    ax3.axis('off')
    ax3.set_title('SL', fontsize=10)
    ax4.set_xlim(-2, 2)
    ax4.set_ylim(-0.5, 4.5)
    ax4.axis('off')
    ax4.set_title('CU/KC', fontsize=10)
    ax5.set_xlim(-2, 2)
    ax5.set_ylim(-0.5, 4.5)
    ax5.axis('off')
    ax5.set_title('CH', fontsize=10)
    ax6.set_xlim(-2, 2)
    ax6.set_ylim(-0.5, 4.5)
    ax6.axis('off')
    ax6.set_title('FS', fontsize=10)
    fig.suptitle('Pitch Locations (Catcher\'s View)')
    fig.subplots_adjust(top=.60, wspace=0.0, bottom=.0, left=.00, right=1)
    memfile = BytesIO()
    plt.savefig(memfile)
    # plt.show() for debug
    return memfile


def getData(data):
    pitch_types = getPitchTypes(data)
    pitches = []
    battedball = []
    for i in range(len(pitch_types)):
        is_pitch = data['pitch_type'] == pitch_types[i]
        selected_data = data[is_pitch]
        label = pitch_types[i]
        count = selected_data['pitch_type'].count()
        # pitch info stuff
        swm = selected_data[selected_data['description'] == 'swinging_strike'].count()[
            'description']
        desiredOutcome = ['swinging_strike', 'swinging_strike_blocked', 'foul',
                          'foul_tip', 'hit_into_play', 'hit_into_play_no_out', 'hit_into_play_score']
        total_swings = selected_data.loc[selected_data['description'].isin(desiredOutcome)].count()[
            'description']
        percentage_used = round(
            (count/data['pitch_type'].count()) * 100, 1)
        avgVelo = round(selected_data['release_speed'].dropna().mean(), 1)
        avgEffVelo = round(selected_data['effective_speed'].dropna().mean(), 1)
        avgSpinRate = round(selected_data['release_spin_rate'].dropna().mean(), 0)
        avgHorzBreak = round(12*selected_data['pfx_x'].dropna().mean(), 1)
        avgVertBreak = round(12*selected_data['pfx_z'].dropna().mean(), 1)
        swstrperc = round(swm/total_swings*100, 1)
        # avgBreakDir = round(m.degrees(m.atan2(avgHorzBreak, avgVertBreak)), 0)
        estwOBA = round(selected_data['estimated_woba_using_speedangle'].dropna().mean(), 3)
        wOBA = round(selected_data['woba_value'].dropna().mean(), 3)

        # batted ball stuff
        bbe = selected_data['launch_speed_angle'].dropna().count()
        weak = round(selected_data[selected_data['launch_speed_angle'] == 1].count()[
            'launch_speed_angle']/bbe*100, 1)
        topped = round(selected_data[selected_data['launch_speed_angle'] == 2].count()[
            'launch_speed_angle']/bbe*100, 1)
        under = round(selected_data[selected_data['launch_speed_angle'] == 3].count()[
            'launch_speed_angle']/bbe*100, 1)
        flare = round(selected_data[selected_data['launch_speed_angle'] == 4].count()[
            'launch_speed_angle']/bbe*100, 1)
        solid = round(selected_data[selected_data['launch_speed_angle'] == 5].count()[
            'launch_speed_angle']/bbe*100, 1)
        barrels = round(selected_data[selected_data['launch_speed_angle'] == 6].count()[
            'launch_speed_angle']/bbe*100, 1)
        hardhit = round(selected_data[selected_data['launch_speed'] >= 95].count()[
            'launch_speed_angle']/bbe*100, 1)

        pitch = [label, percentage_used, avgVelo, avgSpinRate,
                 avgHorzBreak, avgVertBreak, swstrperc, estwOBA, wOBA]
        bbs = [label, percentage_used, weak, topped, under, flare, solid,
               barrels, hardhit]
        pitches.append(pitch)
        battedball.append(bbs)
    pitches = pd.DataFrame(pitches, columns=['Pitch Type', '% Thrown', 'Velocity',
                                             'Spin Rate', 'Horizontal Break',
                                             'Vertical Break', 'Whiff Rate', 'xwOBA',
                                             'wOBA'])
    battedball = pd.DataFrame(battedball, columns=['Pitch Type', '% Thrown',
                                                   'Weak%', 'Topped%', 'Under%',
                                                   'Burner%', 'Solid%',
                                                   'Barrel%', 'HardHit%'])
    pitches = pitches[pitches['% Thrown'] >= 1.5]
    battedball = battedball[battedball['% Thrown'] >= 1.5]
    pitches = pitches.sort_values(by=['% Thrown'], ascending=False)
    battedball = battedball.sort_values(by=['% Thrown'], ascending=False)

    return pitches, battedball


def GenerateProfile(fname, lname, date1, date2, moves, locs, pitches, battedballs, count):
    document = Document()

    paragraph = document.add_paragraph()
    paragraph_format = paragraph.paragraph_format
    paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    titlestr = 'Pitch Profile for ' + fname + ' ' + lname
    datestr = '\n' + date1 + ' to ' + date2
    run0 = paragraph.add_run(titlestr)
    run1 = paragraph.add_run(datestr)
    font0 = run0.font
    font0.size = Pt(24)
    font0.bold = True
    font0.name = 'Calibri'
    font1 = run1.font
    font1.size = Pt(12)
    font1.name = 'Calibri'

    move_paragraph = document.add_paragraph()
    paragraph_format = move_paragraph.paragraph_format
    paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = move_paragraph.add_run()
    my_image = run.add_picture(moves)

    loc_paragraph = document.add_paragraph()
    paragraph_format = loc_paragraph.paragraph_format
    paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = loc_paragraph.add_run()
    my_image = run.add_picture(locs)

    # First row are table headers
    pitchTable = document.add_table(
        pitches.shape[0]+1, pitches.shape[1], style='MediumList2')

    # add the header rows.
    for j in range(pitches.shape[-1]):
        pitchTable.cell(0, j).text = pitches.columns[j]

    # add the rest of the data frame
    for i in range(pitches.shape[0]):
        for j in range(pitches.shape[-1]):
            pitchTable.cell(i+1, j).text = str(pitches.values[i, j])

    break_paragraph = document.add_paragraph('')

    # First row are table headers
    bbTable = document.add_table(
        battedballs.shape[0]+1, battedballs.shape[1], style='MediumList2')

    # add the header rows.
    for j in range(battedballs.shape[-1]):
        bbTable.cell(0, j).text = battedballs.columns[j]

    # add the rest of the data frame
    for i in range(battedballs.shape[0]):
        for j in range(battedballs.shape[-1]):
            bbTable.cell(i+1, j).text = str(battedballs.values[i, j])

    document.save(fname + lname + date2 + '.docx')
    moves.close()
    locs.close()


def main():
    while (1 == 1):
        fname = input('Enter First Name: ')
        lname = input('Enter Last Name: ')
        date1 = input('Enter Start of Date Range (YYYY-MM-DD): ')
        date2 = input('Enter End of Date Range (YYYY-MM-DD): ')
        data = dataGrab(getNumber(lname, fname), date1, date2)
        pitchdata, battedballdata = getData(data)
        releasemovement = plotData(data)
        locations = plotLocationData(data)
        pitches = getPitchTypes(data)
        count = 0
        for i in range(len(pitches)):
            count += 1
        GenerateProfile(fname, lname, date1, date2, releasemovement,
                        locations, pitchdata, battedballdata, count)
        unused_variable = os.system("cls")
        print('Pitch Profile for: ' + fname + lname + date2 + '.docx Created!')

        dec = input('Enter New Pitcher (Y/N)?: ').upper()
        if(dec == 'N'):
            break
        else:
            continue


if __name__ == '__main__':
    main()
