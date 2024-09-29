import os
import os.path as pa
import sys
import TkEasyGUI as sg
import io
import shutil
import subprocess
from zipfile import ZipFile

# Emu = 'd:/Emu/emuZ700/mz700.exe'  # 起動時テープセット可、ロード音OFF推奨
# Emu = 'd:/Emu/mz700/mz700win058.exe'  # 起動時テープセット不可
# Emu = ''  # playボタンなし

Emu_list = {'emuZ700': 'd:/Emu/emuZ700/mz700.exe',
            'MZ700Win 058': 'd:/Emu/mz700/mz700win058.exe'
            }


Emu = 'emuZ700'  # 起動時テープセット可、ロード音OFF推奨
Conv = './MZTMERGE.EXE'
Nheader = 'NHEADER.BIN'

Config = 'temp.cfg'
DfltGba = 'MZ700GBA.GBA'

Cfg_core = 'M7GBA_GP.BIN'
Cfg_core_opt = ['M7GBA_GP.BIN', 'M7GBA_JC.BIN']
#   CORE:	エミュレータコアのファイル名を指定します。
#		M7GBA_GP.BIN	ゲームパック用（08000000h）
#		M7GBA_JC.BIN	ヒカルの碁3ジョイキャリーカートリッジ用（09F00000h）

Cfg_iplrom = '1Z-009A.ROM'
Cfg_ipl_opt = ['1Z-009A.ROM', 'NEWMON7.ROM', '1Z-009B7.ROM', 'SP-1002.ROM']
#   IPLROM:	IPL ROM イメージ(1Z-009A)のファイル名を指定します。
#

Cfg_title = '*'
#  TITLE:	テープイメージのタイトル。16字までの英数字
Cfg_tapeimg = 'TEMP.MZT'
#   TAPEIMG:	テープイメージのファイル名

Cfg_key_a = 'CTRL'
Cfg_key_b = 'SPACE'
Cfg_key_l = 'SPACE'
Cfg_key_r = 'CTRL'
Cfg_key_start = 'CR'
#   A:		GBAのAボタンに割り当てるMZ-700のキー
#   B:		GBAのBボタンに割り当てるMZ-700のキー
#   L:		GBAのLボタンに割り当てるMZ-700のキー
#   R:		GBAのRボタンに割り当てるMZ-700のキー
#   START:	GBAのSTARTボタンに割り当てるMZ-700のキー

Cfg_keys = ['',
            'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M',
           'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
           '0', '1', '2', '3', '4', '5', '6', '7', '8', '9',
           '(', ')', '*', '+', '-', '=', '@', ';', ':', ',', '.', '?', '/',
           'SPACE', 'CR', 'CTRL', 'SHIFT', 'BREAK',
           'KANA', 'GRAPH', 'EISU', 'INS', 'DEL', 'F1', 'F2', 'F3', 'F4', 'F5'
           ]
#   キーに記述できる文字列

Hubasic_auto = [ 0x50, 0x41, 0x54, 0x3A, 0x3B, 0x12, 0x0C, 0xED,
                 0xF4, 0x07, 0xED, 0xF4, 0x08, 0xED, 0xF4, 0x05,
                 0xED, 0xF4, 0x06, 0x7A, 0x26, 0x06, 0x05, 0x52,
                 0x55, 0x4E, 0x22, 0x0D, 0xFF, 0xFF, 0x41, 0x4B,
                 0x50, 0x3A, 0x09, 0x80, 0x3C, 0x78, 0x00, 0x00,
                 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00
                 ]
#Hu-BASICプログラムの自動起動
#   Hu-BASICのMZTファイルのヘッダのオフセット40hからを上記に書き換えると、
#   MZT上のHu-BASICの直後にあるBASICプログラムを自動で実行できます。

def lo_config_input():
    keys, lo_fn = lo_fname()
    k, lo_co = lo_cores()
    keys.extend(k)
    k, lo_ks = lo_keyset()
    keys.extend(k)
    k, lo_bt = lo_butt()
    keys.extend(k)

    return keys, [[lo_fn], [lo_co, lo_ks, lo_bt],
                  [sg.Multiline('', key='-buf-',
                                size=(80,20), readonly=True,
                                readonly_background_color='gray25',
                                expand_x=True, expand_y=True)]]


def lo_fname():
    title = [sg.Text('Title', width=11),
              sg.Input(key='-title-', default_text='', width=41)]
    fname = sg.Text('', key='-fname-', size=(61,2), wrap_length=400,
                     background_color='#ffffff', expand_x=True)    
    typs = (('Tape Image', '*.mzt *.m12'),
            ('Zip File', '*.zip'),
            ('All Files', '*.*'),)
    fbtn = sg.FileBrowse('Image', target_key='-fname-',
                         file_types=typs)
    file = [sg.Text('File Name', width=11),
            fname, fbtn]
    keys = ['-title-', '-fname-']

    return keys, sg.Column(layout=[title, file],
                           expand_x=True, pad=0,
                           key='fname')


def lo_keyset():
    keys_lb = sg.Column([[sg.Text('L'),
                          sg.Combo(values=Cfg_keys, default_value='SPACE',
                                   width=6, key='-keyl-')],
                         [sg.Text('B'),
                          sg.Combo(values=Cfg_keys, default_value='SPACE',
                                   width=6, key='-keyb-')]
                         ])
    keys_ra = sg.Column([[sg.Text('R'),
                          sg.Combo(values=Cfg_keys, default_value='CTRL',
                                   width=6, key='-keyr-')],
                         [sg.Text('A'),
                          sg.Combo(values=Cfg_keys, default_value='CTRL',
                                   width=6, key='-keya-')]
                         ])
    key_start = sg.Combo(values=Cfg_keys, default_value='CR',
                         width=6, key='-keys-')
    keys = ['-keya-', '-keyb-', '-keyl-', '-keyr-', '-keys-']

    return keys, sg.Frame('Key Binds',
                          layout=[[keys_lb, keys_ra,
                                   sg.Text('START'), key_start]],
                          label_outside=True, key='keyset')


def lo_cores():
    labels = sg.Column([[sg.Text('Core')],[sg.Text('IPL ')]])
    settings = sg.Column([[sg.Combo(values=Cfg_core_opt,
                                   default_value=Cfg_core, key='-core-')],
                          [sg.Combo(values=Cfg_ipl_opt,
                                    default_value=Cfg_iplrom, key='-ipl-')]])
    
    keys=['-core-', '-ipl-']
    return keys, sg.Frame('Core Settings',
                          layout=[[labels, settings]],
                          label_outside=True, key='cores')


def lo_butt():
    conf = sg.Button('Show Config', key='-test-')
    exec = sg.Button('Create',
                     color='#ffffff', background_color='#0088aa',
                     key='-go-')

    sel = sg.Combo(values=list(Emu_list.keys()), default_value=Emu,
                   width=8, enable_events=True, key='-emu-')
    hdrcpy = sg.Checkbox('N Header', default=True, key='-nin-')
    play = sg.Button('Play', key='-play-')
    
    keys = ['-test-', '-go-', '-play-', '-emu-', '-nin-']

    return keys, sg.Column([[sg.Text('', expand_x=True),
                             hdrcpy, exec],
                            [sg.Text('', expand_x=True),
                             conf],
                            [sg.Text('', expand_x=True),
                             sel, play]],
                           vertical_alignment='bottom',
                           expand_x=True)

    
Cfg_key = {'CORE': '-core-',
           'IPLROM': '-ipl-',
           'TITLE': '-title-',
            'TAPEIMG': 'temp.mzt',
            'A': '-keya-',
            'B': '-keyb-',
            'L': '-keyl-',
            'R': '-keyr-',
            'START': '-keys-'
            }


def make_config(wn, keys):
    buf = io.StringIO()
    
    ttl = wn['-title-'].get()
    fnam = wn['-fname-'].get()

    print('FNAME: %s (%2d) '%(fnam,len(fnam)),
          'TITLE: %s (%2d)\n'%(ttl,len(ttl)))
    if len(fnam) < 1:
        return ''
    
    if len(ttl) < 1:
        ttl = cleanup(pa.basename(fnam))
    if not ttl.isalnum():
        ttl = cleanup(ttl)
    ttl = ttl.upper()[:16]
    wn['-title-'].update(text=ttl)

    for k in Cfg_key:
        if k == 'TAPEIMG':
            sprintf(buf, 'TAPEIMG:\t%s\n' % Cfg_tapeimg)  # 制限があるので固定
        elif Cfg_key[k] in keys:
            v = str(wn[Cfg_key[k]].get())
            if v != '':
                print('----', k, Cfg_key[k], v)
                sprintf(buf, '%s:\t%s\n'%(k, v))
    
    return buf.getvalue()


def prep_cpy(wn, cnfs):
    src = wn['-fname-'].get()
    if src == '':
        return False

    if pa.isfile(Cfg_tapeimg):
        os.remove(Cfg_tapeimg)

    if pa.splitext(src)[1].upper() == '.ZIP':
        if mzf_in_zip(src, wn) is None:
            return False
    else:
        try:
            shutil.copy(src, Cfg_tapeimg)
        except FileNotFoundError:
            return False
    
    with open(Config, mode='w+', encoding='sjis') as f:
        f.writelines(cnfs)

    return True


def exec_merge(wn):
    command = [Conv, Config]
    res = subprocess.run(command, capture_output=True, text=True)
    out = res.stdout
    if (out is None) or (out == ''):
        out = res.stderr
    if (out is not None) and (len(out) > 0):
        wn['-buf-'].update(text=out)
        wn.refresh()
        return False
    
    tgtfile = 'MZ7_'+wn['-title-'].get()+'.GBA'
    tgt = open(tgtfile, mode='wb')
    
    if wn['-nin-'].get() == True:
        # 任天堂ヘッダ書込み
        hdr = open(Nheader, mode='rb')
        data = hdr.read(192)
        tgt.write(data)
        hdr.close()
        seeksize = 192
    else:
        seeksize = 0

    # 本体+ROM書込み
    gba = open(DfltGba, mode='rb')
    gba.seek(seeksize)
    data = gba.read()
    tgt.write(data)
    gba.close()

    tgt.close()    
    return True


def mzf_in_zip(zfname, wn):
    with ZipFile(zfname) as zf:
        l = zf.filelist

    tb = []
    count = 0
    target = ''
    for x in l:
        nm = x.filename
        if ('.MZF' in nm.upper()) or ('.M12' in nm.upper()):
            count += 1
            target = nm
        tb.append([nm, x.file_size])

    if count == 0:
        return None
    elif count > 1:
        # 複数候補の場合、リストから選択
        tbl = sg.Table(values=tb, headings=['Filename','Size'], key='-zt-',
                       event_returns_values=True,
                       select_mode='single', expand_x=True)
        lo = [[tbl],[sg.Text('', expand_x=True),
                     sg.Button('Cancel', key='-zc-'),
                     sg.Button('Open', key='-zo-')]]
        wn.set_alpha_channel(0.3)
        wn2 = sg.Window(zfname, lo)

        while True:
            e,v = wn2.read()
            if e in [sg.WINDOW_CLOSED, '-zo-', '-zc-']:
                break

        # ボタンを押してもリストが選択されてなければ中断
        wn2.close()
        wn.set_alpha_channel(1.0)
        if (v['-zt-'] == '') or (e == '-zc-'):
            return None
        target = v['-zt-'][0]

    if target == '':
        return None

    with ZipFile(zfname) as zf:
        zf.extract(target, path='.')
        os.rename(target, Cfg_tapeimg)

    return Cfg_tapeimg


def emuplay(file):
    absfile = pa.abspath(file)
    emudir, emuexe = pa.split(Emu_list[Emu])
    print('%s %s, cwd=%s'%(emuexe, absfile, emudir))
    
    res = subprocess.run([Emu_list[Emu], absfile], cwd=emudir,
                         capture_output=True, text=True)
    return True


Rmtmp = [DfltGba, Config, Cfg_tapeimg]
def remove_temp():
    for x in Rmtmp:
        if pa.isfile(x):
            os.remove(x)
    return


def sprintf(buf, *args, **kwargs):
    s = []
    for element in args:
        s.append(str(element))
        #' '.join(s)
        
    return buf.write(*args, **kwargs)


def cleanup(s):
    d = ''
    for x in s:
        if (not x.isalnum()) and (x != '_'):
            d += '_'
        else:
            d += x
    return d.upper()
    
if __name__ == '__main__':
    keys, lo = lo_config_input()
    wn = sg.Window('MZT MERGE', layout=lo )

    while True:
        e,v = wn.read()
        print(e, v)

        if e == sg.WINDOW_CLOSED:
            wn.close()
            break
        elif e == '-emu-':
            Emu = wn['-emu-'].get()

        elif e in ['-go-', '-test-', '-play-']:
            cnfs = make_config(wn, keys)
            if cnfs == '':
                wn['-buf-'].update(text='*** BAD CONFIG ***')
                wn.refresh()
                continue
            wn['-buf-'].update(text=cnfs)
            wn.refresh()
            if e == '-test-':
                continue
            if not prep_cpy(wn, cnfs):
                wn['-buf-'].update(text='COPY Failed...')
                wn.refresh()
                remove_temp()
                continue
            if e == '-go-':
                if not exec_merge(wn):
                    remove_temp()
                    continue
                wn['-buf-'].update(text=cnfs+'*** Complete conversion ***\n')
                wn.refresh()
            elif e == '-play-':
                em = wn['-emu-'].get()
                if em != Emu:
                    if em not in Emu_list:
                        wn['-emu-'].update(value=Emu)
                    else:
                        Emu = em
                emuplay(Cfg_tapeimg)
            remove_temp()
