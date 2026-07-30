"""
S 参数可视化 GUI — 支持 dB / RI / CSV / TXT 格式
  - 文件浏览、格式自动识别
  - S11/S21 幅值 + 相位绘图 (单面板 / 双面板 / 2×2 四子图)
  - IEEE TAP 学术风格
  - 支持批量加载、保存图片

用法: python plot_sparameter_gui.py
"""

import os
import re
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

# ═══════════════════════════════════════════════════════════
#  IEEE TAP 全局样式
# ═══════════════════════════════════════════════════════════
mpl.rcParams.update({
    'font.family':        'serif',
    'font.serif':         ['Times New Roman'],
    'mathtext.fontset':   'stix',
    'font.size':           9,
    'axes.titlesize':      9,
    'axes.labelsize':      9,
    'xtick.labelsize':     8,
    'ytick.labelsize':     8,
    'legend.fontsize':     7.5,
    'lines.linewidth':     1.0,
    'axes.linewidth':      0.8,
    'xtick.major.width':   0.8,
    'ytick.major.width':   0.8,
    'grid.linewidth':      0.4,
    'grid.alpha':          0.4,
    'grid.linestyle':      '--',
    'legend.framealpha':   0.85,
    'legend.edgecolor':    '0.8',
    'legend.handlelength': 2.0,
    'figure.dpi':          150,
    'savefig.dpi':         600,
    'figure.facecolor':    'white',
    'axes.facecolor':      'white',
})

# ═══════════════════════════════════════════════════════════
#  常量
# ═══════════════════════════════════════════════════════════
# 红绿渐变色 (小角度→绿 #47BA78, 大角度→红 #EA6759)
COLOR_GREEN = '#47BA78'
COLOR_RED   = '#EA6759'
# 预建 colormap: 绿(0.0) → 红(1.0)，用于角度连续映射
CMAP_RG = mpl.colors.LinearSegmentedColormap.from_list('red_green', [COLOR_GREEN, COLOR_RED])

# 多文件叠加时的区分色系 (每个文件一套 colormap)
FILE_CMAPS = [
    CMAP_RG,                                                          # 绿→红
    mpl.colors.LinearSegmentedColormap.from_list('blue_orange',       # 蓝→橙
        ['#1f77b4', '#ff7f0e']),
    mpl.colors.LinearSegmentedColormap.from_list('purple_brown',      # 紫→棕
        ['#9467bd', '#8c564b']),
    mpl.colors.LinearSegmentedColormap.from_list('cyan_pink',         # 青→粉
        ['#17becf', '#e377c2']),
]

TAP_LS = {'TE': '-', 'TM': '--'}  # 极化线型

FIG_SINGLE  = (3.5, 2.8)   # 单栏
FIG_WIDE    = (7.16, 3.0)  # 双栏
FIG_DOUBLE  = (7.16, 5.5)  # 双栏 2 行
FIG_2X2     = (7.5, 5.2)   # 2×2 四子图


# ═══════════════════════════════════════════════════════════
#  数据解析
# ═══════════════════════════════════════════════════════════

def _detect_csv_format(filepath):
    """检测 CSV 文件格式: 'dB' 或 'RI' 或 None"""
    try:
        df = pd.read_csv(filepath, nrows=0)
        cols = list(df.columns)
        if any('Mag_dB' in c for c in cols):
            return 'dB'
        if any('_Re' in c for c in cols) or any('_Im' in c for c in cols):
            return 'RI'
        return None
    except Exception:
        return None


def parse_csv(filepath):
    """
    解析 CSV 文件，自动识别 dB / RI 格式。
    返回 dict: {
        'freq': np.array (GHz),
        'angle': np.array (deg),
        'data': {(s_type, pol): {'mag_dB': array, 'phase_deg': array}}
    }
    """
    df = pd.read_csv(filepath)
    fmt = _detect_csv_format(filepath)

    # 去重
    dup_cols = ["Angle", "Frequency_GHz"]
    if df.duplicated(subset=dup_cols).any():
        df = df.groupby(dup_cols, as_index=False).mean()

    freq = df["Frequency_GHz"].values
    angles = df["Angle"].values

    data = {}
    s_types = ['S11', 'S21']
    pols = ['TE', 'TM']

    for st in s_types:
        for pol in pols:
            if fmt == 'dB':
                mag_col  = f'{st}_{pol}_Mag_dB'
                phase_col = f'{st}_{pol}_Phase_deg'
                if mag_col in df.columns and phase_col in df.columns:
                    data[(st, pol)] = {
                        'mag_dB': df[mag_col].values,
                        'phase_deg': df[phase_col].values,
                    }
            elif fmt == 'RI':
                re_col = f'{st}_{pol}_Re'
                im_col = f'{st}_{pol}_Im'
                if re_col in df.columns and im_col in df.columns:
                    s_cplx = df[re_col].values + 1j * df[im_col].values
                    data[(st, pol)] = {
                        'mag_dB': 20 * np.log10(np.abs(s_cplx) + 1e-15),
                        'phase_deg': np.angle(s_cplx, deg=True),
                    }
    return {
        'freq': freq,
        'angle': angles,
        'data': data,
        'format': fmt,
        'filename': os.path.basename(filepath),
        'filepath': filepath,
    }


def _is_txt_amplitude(lines):
    """通过 TXT 头行判断是幅值(dB)还是相位(deg)"""
    for line in lines[:3]:
        if 'abs,dB' in line or 'abs, dB' in line or 'Amplitude' in line.lower():
            return True
        if 'arg,degrees' in line or 'arg, degrees' in line or 'phase' in line.lower() or 'Phase' in line:
            return False
    return None


def parse_txt_single(filepath):
    """
    解析单个 TXT 文件 (幅值或相位)。
    返回 dict: {'freq': array, 'value': array, 'type': 'amp'|'phase'}
    自动对重复频率点取均值去重。
    """
    with open(filepath, 'r') as f:
        lines = f.readlines()

    is_amp = _is_txt_amplitude(lines)
    data_type = 'amp' if is_amp else 'phase'

    # Keep the CST theta header associated with every following numeric row.
    data_lines = []
    current_angle = 0.0
    for line in lines:
        line = line.strip()
        if not line or line.startswith('-'):
            continue
        angle_match = re.search(r'\btheta\s*=\s*([-+]?\d*\.?\d+)', line,
                                flags=re.IGNORECASE)
        if angle_match:
            current_angle = float(angle_match.group(1))
            continue
        # 尝试解析为两个浮点数
        parts = line.split()
        if len(parts) >= 2:
            try:
                float(parts[0])
                float(parts[1])
                data_lines.append((float(parts[0]), float(parts[1]), current_angle))
            except ValueError:
                continue

    if not data_lines:
        raise ValueError(f'No numeric CST data found in: {os.path.basename(filepath)}')
    freq = np.array([p[0] for p in data_lines])
    value = np.array([p[1] for p in data_lines])
    angle = np.array([p[2] for p in data_lines])

    # Frequency repeats for each incident angle, so deduplicate by angle + frequency.
    if len(freq) != len(pd.MultiIndex.from_arrays([angle, freq]).unique()):
        df_tmp = pd.DataFrame({'angle': angle, 'freq': freq, 'value': value})
        df_tmp = df_tmp.groupby(['angle', 'freq'], as_index=False).mean()
        freq = df_tmp['freq'].values
        value = df_tmp['value'].values
        angle = df_tmp['angle'].values

    return {
        'freq': freq,
        'value': value,
        'angle': angle,
        'type': data_type,
        'filename': os.path.basename(filepath),
        'filepath': filepath,
    }


def _guess_txt_pair(filepath):
    """
    根据 TXT 文件路径，猜测对应的配对文件。
    - 输入幅值文件 (*_Am.txt) → 返回相位文件 (*_Ph.txt) 路径
    - 输入相位文件 (*_Ph.txt) → 返回幅值文件 (*_Am.txt) 路径
    - 找不到则返回 None
    """
    dirname = os.path.dirname(filepath)
    basename = os.path.basename(filepath)

    # 规则1: _Am.txt ↔ _Ph.txt (双向)
    if '_Am' in basename:
        target = basename.replace('_Am', '_Ph')
        target_path = os.path.join(dirname, target)
        if os.path.exists(target_path):
            return target_path
    elif '_Ph' in basename:
        target = basename.replace('_Ph', '_Am')
        target_path = os.path.join(dirname, target)
        if os.path.exists(target_path):
            return target_path

    # 规则2: 同目录下搜索可能的配对 (尝试常见后缀)
    base_no_ext = os.path.splitext(basename)[0]
    for a_tag, p_tag in [('Am', 'Ph'), ('amp', 'phase'), ('Amplitude', 'Phase')]:
        tag_found = None
        other_tag = None
        if a_tag in base_no_ext:
            tag_found = a_tag
            other_tag = p_tag
        elif p_tag in base_no_ext:
            tag_found = p_tag
            other_tag = a_tag
        if tag_found:
            target_base = base_no_ext.replace(tag_found, other_tag)
            for ext in ['.txt', '.csv', '.dat']:
                candidate = os.path.join(dirname, target_base + ext)
                if os.path.exists(candidate):
                    return candidate
    return None


def parse_txt_paired(filepath):
    """
    解析 TXT 幅值 + 相位配对文件 (输入可以是 Am 或 Ph 文件)。
    返回与 parse_csv 相同结构的 dict。
    """
    first_data = parse_txt_single(filepath)
    pair_path = _guess_txt_pair(filepath)
    if pair_path is None:
        return None

    second_data = parse_txt_single(pair_path)

    # 确定哪个是幅值，哪个是相位
    if first_data['type'] == 'amp':
        amp_data = first_data
        phase_data = second_data
    else:
        amp_data = second_data
        phase_data = first_data

    # 从文件名猜测 S 参数和极化
    basename = os.path.basename(filepath)
    s_type = 'S21'
    if 'S11' in basename:
        s_type = 'S11'
    elif 'S21' in basename:
        s_type = 'S21'

    pol = 'TE'
    if 'TM' in basename:
        pol = 'TM'
    elif 'TE' in basename:
        pol = 'TE'

    paired = pd.merge(
        pd.DataFrame({'Angle': amp_data['angle'], 'Frequency_GHz': amp_data['freq'],
                      'mag_dB': amp_data['value']}),
        pd.DataFrame({'Angle': phase_data['angle'], 'Frequency_GHz': phase_data['freq'],
                      'phase_deg': phase_data['value']}),
        on=['Angle', 'Frequency_GHz'], how='inner')
    if paired.empty:
        return None
    freq = paired['Frequency_GHz'].values

    data = {
        (s_type, pol): {
            'mag_dB': paired['mag_dB'].values,
            'phase_deg': paired['phase_deg'].values,
        }
    }

    return {
        'freq': freq,
        'angle': paired['Angle'].values,
        'data': data,
        'format': 'txt_paired',
        'filename': os.path.basename(filepath) + ' + ' + os.path.basename(pair_path),
        'filepath': filepath,
    }


def convert_txt_folder_to_lcalpha_csv(folder, output_path):
    """Convert CST *_Am.txt / *_Ph.txt exports into LCalpha RI CSV input."""
    expected = ('S11_TE', 'S11_TM', 'S21_TE', 'S21_TM')
    columns = ['Angle', 'Frequency_GHz']
    columns += [f'{name}_{part}' for name in expected for part in ('Re', 'Im')]

    pairs = {}
    for filename in os.listdir(folder):
        # LCalpha uses co-polarized S parameters. Ignore optional *_Cross_* exports.
        if re.search(r'_Cross_', filename, flags=re.IGNORECASE):
            continue
        match = re.search(r'(?:((?:S(?:11|21))_(?:TE|TM))|((?:TE|TM)_(?:S(?:11|21))))_(Am|Ph)\.txt$',
                          filename, flags=re.IGNORECASE)
        if match:
            parameter = (match.group(1) or match.group(2)).upper()
            first, second = parameter.split('_')
            name = parameter if first.startswith('S') else f'{second}_{first}'
            kind = match.group(3).lower()
            pairs.setdefault(name, {})[kind] = os.path.join(folder, filename)

    missing = [name for name in expected
               if not {'am', 'ph'} <= set(pairs.get(name, {}))]
    if missing:
        raise ValueError('Missing amplitude/phase TXT pair(s): ' + ', '.join(missing))

    result = None
    for name in expected:
        amplitude = parse_txt_single(pairs[name]['am'])
        phase = parse_txt_single(pairs[name]['ph'])
        amp_df = pd.DataFrame({
            'Angle': amplitude['angle'], 'Frequency_GHz': amplitude['freq'],
            'mag_dB': amplitude['value']})
        phase_df = pd.DataFrame({
            'Angle': phase['angle'], 'Frequency_GHz': phase['freq'],
            'phase_deg': phase['value']})
        parameter = pd.merge(amp_df, phase_df,
                             on=['Angle', 'Frequency_GHz'], how='inner')
        if parameter.empty:
            raise ValueError(f'No matching angle/frequency points for {name}')
        magnitude = np.power(10.0, parameter.pop('mag_dB') / 20.0)
        phase_rad = np.deg2rad(parameter.pop('phase_deg'))
        values = magnitude * np.exp(1j * phase_rad)
        parameter[f'{name}_Re'] = np.real(values)
        parameter[f'{name}_Im'] = np.imag(values)
        result = (parameter if result is None else
                  pd.merge(result, parameter,
                           on=['Angle', 'Frequency_GHz'], how='inner'))

    if result.empty:
        raise ValueError('The four S-parameter pairs have no common angle/frequency points.')
    result = result.sort_values(['Angle', 'Frequency_GHz']).reset_index(drop=True)

    result.to_csv(output_path, index=False, columns=columns)
    return len(result)


def _detect_format(filepath):
    """自动识别文件格式，返回明确的格式标签"""
    ext = os.path.splitext(filepath)[1].lower()
    if ext == '.csv':
        fmt = _detect_csv_format(filepath)
        return f'csv_{fmt}' if fmt else 'unknown'
    elif ext in ('.txt', '.dat'):
        # 检查是否为相位文件 (本身不加载，由配对逻辑处理)
        is_phase_file = 'Ph' in os.path.basename(filepath)
        has_pair = bool(_guess_txt_pair(filepath))

        if is_phase_file and not has_pair:
            return 'txt_phase_unpaired'  # 独立相位文件

        if has_pair or is_phase_file:
            return 'txt_paired'

        # 无配对的独立 TXT
        try:
            with open(filepath, 'r') as f:
                lines = [f.readline() for _ in range(3)]
            amp = _is_txt_amplitude(lines)
            if amp is True:
                return 'txt_amp'
            elif amp is False:
                return 'txt_phase'
        except Exception:
            pass
        return 'unknown'
    return 'unknown'


def load_file(filepath, fmt='auto'):
    """
    统一的文件加载接口。
    fmt: 'auto' | 'csv_dB' | 'csv_RI' | 'txt_paired' | 'txt_amp' | 'txt_phase'
    返回 parsed dict 或 None
    """
    if fmt == 'auto':
        fmt = _detect_format(filepath)

    try:
        if fmt in ('csv_dB', 'csv_RI'):
            return parse_csv(filepath)
        elif fmt == 'txt_paired':
            return parse_txt_paired(filepath)
        elif fmt in ('txt_amp', 'txt_phase'):
            return parse_txt_single(filepath)
        else:
            # fallback: 尝试所有格式
            result = parse_csv(filepath)
            if result and result.get('data'):
                return result
            result = parse_txt_paired(filepath)
            if result:
                return result
            return parse_txt_single(filepath)
    except Exception:
        return None


# ═══════════════════════════════════════════════════════════
#  辅助函数
# ═══════════════════════════════════════════════════════════

def _filter_angles(all_angles, angles):
    """过滤并返回要绘制的角度列表"""
    if angles is None or (isinstance(angles, str) and angles.strip() == ''):
        return sorted(all_angles)
    if isinstance(angles, (int, float, np.integer, np.floating)):
        angles = [angles]
    if isinstance(angles, str):
        angles = [float(a.strip()) for a in angles.split(',') if a.strip()]
    return sorted([a for a in angles if a in all_angles])


def _get_color(angle, plot_angles, cmap=None):
    """角度→颜色映射。cmap 指定色系，默认红绿渐变。"""
    if cmap is None:
        cmap = CMAP_RG
    if len(plot_angles) <= 1:
        return mpl.colors.to_hex(cmap(0.0))
    amin, amax = min(plot_angles), max(plot_angles)
    t = (angle - amin) / (amax - amin) if amax != amin else 0.0
    return mpl.colors.to_hex(cmap(t))


def _angle_label(ang):
    """格式化角度标签"""
    return f'{ang:.0f}$^\\circ$' if ang == int(ang) else f'{ang:.1f}$^\\circ$'


# ═══════════════════════════════════════════════════════════
#  绘图函数 (接收 parsed_data dict)
# ═══════════════════════════════════════════════════════════

def _select_data(parsed, s_type, pols, angles, freq_range):
    """
    从 parsed dict 中提取指定参数的数据。
    返回 [(pol, ang, freq, mag_dB, phase_deg), ...]
    """
    freq = parsed['freq']
    angle_arr = parsed['angle']
    data_dict = parsed['data']

    all_angles = np.unique(angle_arr)
    plot_angles = _filter_angles(all_angles, angles)

    results = []
    for pol in pols:
        key = (s_type, pol)
        if key not in data_dict:
            continue
        mag = data_dict[key]['mag_dB']
        ph  = data_dict[key]['phase_deg']

        for ang in plot_angles:
            # 对于 CSV 格式: 有 angle 列，筛选角度
            # 对于 TXT 格式: angle_arr 全为 0，直接用全部
            if len(np.unique(angle_arr)) > 1:
                mask = angle_arr == ang
                f = freq[mask]
                m = mag[mask]
                p = ph[mask]
            else:
                f = freq
                m = mag
                p = ph

            # 频率过滤
            if freq_range:
                f_mask = (f >= freq_range[0]) & (f <= freq_range[1])
                f = f[f_mask]
                m = m[f_mask]
                p = p[f_mask]

            # 按频率排序
            order = np.argsort(f)
            results.append((pol, ang, f[order], m[order], p[order]))

    return results


def plot_s_magnitude(parsed, s_type="S21", angles=None, pols=("TE", "TM"),
                     freq_range=None, ylim=(-50, 5), title="", fig=None, ax=None,
                     decorate=True, cmap=None):
    """仅绘制 |S| 幅值 (dB)，单面板。decorate=False 时只画数据线不设标签/图例/参考线。"""
    lines_data = _select_data(parsed, s_type, pols, angles, freq_range)
    if not lines_data:
        return None

    if fig is None:
        fig, ax = plt.subplots(figsize=FIG_SINGLE)

    all_plot_angles = sorted(set(ang for _, ang, _, _, _ in lines_data))
    for pol, ang, f, mag, _ in lines_data:
        c = _get_color(ang, all_plot_angles, cmap=cmap)
        ax.plot(f, mag, color=c, linestyle=TAP_LS.get(pol, '-'),
                linewidth=1.0, marker='',
                label=f'{pol} {_angle_label(ang)}')

    if decorate:
        ax.set_xlabel("Frequency (GHz)")
        ax.set_ylabel(f"|{s_type}| (dB)")
        ax.set_ylim(ylim)
        if freq_range:
            ax.set_xlim(freq_range)
        ax.axhline(-3,  color='gray', linestyle=':', linewidth=0.6, alpha=0.5)
        ax.axhline(-10, color='gray', linestyle=':', linewidth=0.6, alpha=0.5)
        ax.grid(True)
        ax.legend(loc='lower left', ncol=2, fontsize=7)
        if title:
            ax.set_title(title, fontsize=9)
        fig.tight_layout()
    return fig


def plot_s_phase(parsed, s_type="S21", angles=None, pols=("TE", "TM"),
                 freq_range=None, ylim=(-200, 200), title="", fig=None, ax=None,
                 decorate=True, cmap=None):
    """仅绘制相位 (deg)，单面板。decorate=False 时只画数据线。"""
    lines_data = _select_data(parsed, s_type, pols, angles, freq_range)
    if not lines_data:
        return None

    if fig is None:
        fig, ax = plt.subplots(figsize=FIG_SINGLE)

    all_plot_angles = sorted(set(ang for _, ang, _, _, _ in lines_data))
    for pol, ang, f, _, ph in lines_data:
        c = _get_color(ang, all_plot_angles, cmap=cmap)
        ax.plot(f, ph, color=c, linestyle=TAP_LS.get(pol, '-'),
                linewidth=1.0, marker='',
                label=f'{pol} {_angle_label(ang)}')

    if decorate:
        ax.set_xlabel("Frequency (GHz)")
        ax.set_ylabel(f"Phase({s_type}) (deg)")
        ax.set_ylim(ylim)
        if freq_range:
            ax.set_xlim(freq_range)
        ax.axhline(0, color='gray', linestyle=':', linewidth=0.6, alpha=0.5)
        ax.grid(True)
        ax.legend(loc='lower left', ncol=2, fontsize=7)
        if title:
            ax.set_title(title, fontsize=9)
        fig.tight_layout()
    return fig


def plot_s_mag_and_phase(parsed, s_type="S21", angles=None, pols=("TE", "TM"),
                         freq_range=None, ylim_mag=(-50, 5), ylim_phase=(-200, 200),
                         title="", fig=None, axes=None, decorate=True, cmap=None):
    """幅值 + 相位 2×1 双面板。decorate=False 时只画数据线。"""
    lines_data = _select_data(parsed, s_type, pols, angles, freq_range)
    if not lines_data:
        return None

    if fig is None:
        fig, (ax_mag, ax_pha) = plt.subplots(2, 1, figsize=FIG_DOUBLE)
    else:
        ax_mag, ax_pha = axes

    all_plot_angles = sorted(set(ang for _, ang, _, _, _ in lines_data))
    for pol, ang, f, mag, ph in lines_data:
        c = _get_color(ang, all_plot_angles, cmap=cmap)
        lbl = f'{pol} {_angle_label(ang)}'
        ax_mag.plot(f, mag, color=c, linestyle=TAP_LS.get(pol, '-'),
                    linewidth=1.0, marker='', label=lbl)
        ax_pha.plot(f, ph, color=c, linestyle=TAP_LS.get(pol, '-'),
                    linewidth=1.0, marker='', label=lbl)

    if decorate:
        # 幅值面板
        ax_mag.set_ylabel(f"|{s_type}| (dB)")
        ax_mag.set_ylim(ylim_mag)
        if freq_range:
            ax_mag.set_xlim(freq_range)
        ax_mag.axhline(-3,  color='gray', linestyle=':', linewidth=0.6, alpha=0.5)
        ax_mag.axhline(-10, color='gray', linestyle=':', linewidth=0.6, alpha=0.5)
        ax_mag.grid(True)
        ax_mag.legend(loc='lower left', ncol=3, fontsize=6.5)

        # 相位面板
        ax_pha.set_xlabel("Frequency (GHz)")
        ax_pha.set_ylabel(f"Phase({s_type}) (deg)")
        ax_pha.set_ylim(ylim_phase)
        if freq_range:
            ax_pha.set_xlim(freq_range)
        ax_pha.axhline(0, color='gray', linestyle=':', linewidth=0.6, alpha=0.5)
        ax_pha.grid(True)

        if title:
            fig.suptitle(title, fontsize=9, fontweight='bold', y=0.995)
        fig.tight_layout()
    return fig


def plot_s_2x2(parsed, s_type="S21", angles=None, pols=("TE", "TM"),
               freq_range=None, ylim_mag=(-50, 5), ylim_phase=(-200, 200),
               title="", fig=None, axes=None, decorate=True):
    """TE/TM 幅值 + 相位 2×2 四子图，蓝红渐变配色。decorate=False 时只画数据线。"""
    all_angles_raw = np.unique(parsed['angle'])
    plot_angles = _filter_angles(all_angles_raw, angles)
    if not plot_angles:
        return None

    # 蓝红渐变色映射
    angle_min, angle_max = min(plot_angles), max(plot_angles)
    if angle_min == angle_max:
        angle_min, angle_max = angle_min - 1, angle_max + 1
    norm = mpl.colors.Normalize(vmin=angle_min, vmax=angle_max)
    cmap = CMAP_RG

    pol_list = [p for p in ("TE", "TM") if p in pols]
    if not pol_list:
        return None

    if fig is None:
        fig, axs = plt.subplots(2, len(pol_list), figsize=FIG_2X2, squeeze=False)
    else:
        axs = axes

    if decorate and title:
        fig.suptitle(f"{s_type} — {title}", fontsize=10, fontweight='bold')

    for col_idx, pol in enumerate(pol_list):
        ax_mag = axs[0][col_idx]
        ax_pha = axs[1][col_idx]

        lines_data = _select_data(parsed, s_type, [pol], plot_angles, freq_range)
        for _, ang, f, mag, ph in lines_data:
            c = cmap(norm(ang))
            lbl = f'$\\theta={ang:.0f}^\\circ$'
            ax_mag.plot(f, mag, color=c, linestyle='-', linewidth=1.0, label=lbl)
            ax_pha.plot(f, ph, color=c, linestyle='-', linewidth=1.0, label=lbl)

        if decorate:
            # 幅值面板
            ax_mag.set_title(f"{pol} — Amplitude", fontsize=9, fontweight='bold')
            ax_mag.set_ylabel(f"|{s_type}| (dB)")
            ax_mag.set_ylim(ylim_mag)
            if freq_range:
                ax_mag.set_xlim(freq_range)
            ax_mag.axhline(-3,  color='gray', linestyle=':', linewidth=0.6, alpha=0.5)
            ax_mag.axhline(-10, color='gray', linestyle=':', linewidth=0.6, alpha=0.5)
            ax_mag.grid(True)
            ax_mag.legend(loc='lower left', fontsize=7, ncol=2, framealpha=0.85)

            # 相位面板
            ax_pha.set_title(f"{pol} — Phase", fontsize=9, fontweight='bold')
            ax_pha.set_xlabel("Frequency (GHz)")
            ax_pha.set_ylabel(f"Phase({s_type}) (deg)")
            ax_pha.set_ylim(ylim_phase)
            if freq_range:
                ax_pha.set_xlim(freq_range)
            ax_pha.axhline(0, color='gray', linestyle=':', linewidth=0.6, alpha=0.5)
            ax_pha.grid(True)

    if decorate:
        fig.tight_layout(rect=[0, 0, 1, 0.95])
    return fig


# ═══════════════════════════════════════════════════════════
#  GUI 主窗口
# ═══════════════════════════════════════════════════════════

class SParameterGUI:
    """S 参数可视化主窗口"""

    def __init__(self, root):
        self.root = root
        self.root.title("SDC-ECM — S-Parameter Visualization & LC Extraction")
        self.root.geometry("1100x820")
        self.root.minsize(900, 650)

        # 数据存储
        self.loaded_files = {}     # {display_name: parsed_dict}
        self.current_figure = None

        self._build_ui()

    def _build_ui(self):
        """构建完整的 UI 布局"""
        # ── 顶部菜单栏 ──
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Add Files...", command=self._on_add_files)
        file_menu.add_command(label="Add Folder...", command=self._on_add_folder)
        file_menu.add_separator()
        file_menu.add_command(label="Clear All", command=self._on_clear)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        menubar.add_cascade(label="File", menu=file_menu)
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=self._on_about)
        menubar.add_cascade(label="Help", menu=help_menu)
        self.root.config(menu=menubar)

        # ── 主布局: 左侧面板 + 右侧画布 ──
        main_pw = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_pw.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 左侧控制面板
        left_frame = ttk.Frame(main_pw, width=320)
        main_pw.add(left_frame, weight=0)

        # 右侧画布
        right_frame = ttk.Frame(main_pw)
        main_pw.add(right_frame, weight=1)

        # ── 构建各子面板 ──
        self._build_file_panel(left_frame)
        self._build_param_panel(left_frame)
        self._build_plot_controls(left_frame)
        self._build_canvas(right_frame)
        self._build_statusbar()

    # ── 文件面板 ──
    def _build_file_panel(self, parent):
        frame = ttk.LabelFrame(parent, text="File Selection", padding=5)
        frame.pack(fill=tk.BOTH, expand=False, padx=2, pady=2)

        # Listbox + scrollbar
        list_frame = ttk.Frame(frame)
        list_frame.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.file_listbox = tk.Listbox(list_frame, height=6,
                                       yscrollcommand=scrollbar.set,
                                       selectmode=tk.EXTENDED)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.file_listbox.yview)

        # 按钮行
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=3)
        ttk.Button(btn_frame, text="Add Files...", command=self._on_add_files).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Add Folder...", command=self._on_add_folder).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="TXT -> LC CSV...", command=self._on_convert_txt_to_lcalpha_csv).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Remove", command=self._on_remove_file).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Clear", command=self._on_clear).pack(side=tk.LEFT, padx=2)

        # 格式强制选择
        fmt_frame = ttk.Frame(frame)
        fmt_frame.pack(fill=tk.X, pady=3)
        ttk.Label(fmt_frame, text="Format:").pack(side=tk.LEFT, padx=2)
        self.fmt_var = tk.StringVar(value="auto")
        ttk.Radiobutton(fmt_frame, text="Auto", variable=self.fmt_var, value="auto").pack(side=tk.LEFT, padx=1)
        ttk.Radiobutton(fmt_frame, text="CSV-dB", variable=self.fmt_var, value="csv_dB").pack(side=tk.LEFT, padx=1)
        ttk.Radiobutton(fmt_frame, text="CSV-RI", variable=self.fmt_var, value="csv_RI").pack(side=tk.LEFT, padx=1)
        ttk.Radiobutton(fmt_frame, text="TXT Pair", variable=self.fmt_var, value="txt_paired").pack(side=tk.LEFT, padx=1)

    # ── 参数面板 ──
    def _build_param_panel(self, parent):
        frame = ttk.LabelFrame(parent, text="Parameters", padding=5)
        frame.pack(fill=tk.X, padx=2, pady=2)

        # S-Parameter
        row1 = ttk.Frame(frame)
        row1.pack(fill=tk.X, pady=2)
        ttk.Label(row1, text="S-Param:").pack(side=tk.LEFT, padx=2)
        self.stype_var = tk.StringVar(value="S21")
        ttk.Radiobutton(row1, text="S11", variable=self.stype_var, value="S11").pack(side=tk.LEFT, padx=3)
        ttk.Radiobutton(row1, text="S21", variable=self.stype_var, value="S21").pack(side=tk.LEFT, padx=3)

        # Polarization
        row2 = ttk.Frame(frame)
        row2.pack(fill=tk.X, pady=2)
        ttk.Label(row2, text="Pol:").pack(side=tk.LEFT, padx=2)
        self.pol_te = tk.BooleanVar(value=True)
        self.pol_tm = tk.BooleanVar(value=True)
        ttk.Checkbutton(row2, text="TE", variable=self.pol_te).pack(side=tk.LEFT, padx=3)
        ttk.Checkbutton(row2, text="TM", variable=self.pol_tm).pack(side=tk.LEFT, padx=3)

        # Angles
        row3 = ttk.Frame(frame)
        row3.pack(fill=tk.X, pady=2)
        ttk.Label(row3, text="Angles:").pack(side=tk.LEFT, padx=2)
        self.angles_var = tk.StringVar(value="")
        self.angles_entry = ttk.Entry(row3, textvariable=self.angles_var, width=18)
        self.angles_entry.pack(side=tk.LEFT, padx=2)
        ttk.Label(row3, text="(blank=all, comma-sep)", font=('', 7)).pack(side=tk.LEFT)

        # Frequency range
        row4 = ttk.Frame(frame)
        row4.pack(fill=tk.X, pady=2)
        ttk.Label(row4, text="Freq (GHz):").pack(side=tk.LEFT, padx=2)
        self.fmin_var = tk.StringVar(value="")
        self.fmax_var = tk.StringVar(value="")
        ttk.Entry(row4, textvariable=self.fmin_var, width=7).pack(side=tk.LEFT, padx=1)
        ttk.Label(row4, text="–").pack(side=tk.LEFT)
        ttk.Entry(row4, textvariable=self.fmax_var, width=7).pack(side=tk.LEFT, padx=1)
        ttk.Label(row4, text="(blank=all)", font=('', 7)).pack(side=tk.LEFT, padx=3)

        # Y-axis limits
        row5 = ttk.Frame(frame)
        row5.pack(fill=tk.X, pady=2)
        ttk.Label(row5, text="Y-Mag:").pack(side=tk.LEFT, padx=2)
        self.ymin_mag = tk.StringVar(value="-50")
        self.ymax_mag = tk.StringVar(value="5")
        ttk.Entry(row5, textvariable=self.ymin_mag, width=6).pack(side=tk.LEFT, padx=1)
        ttk.Label(row5, text="–").pack(side=tk.LEFT)
        ttk.Entry(row5, textvariable=self.ymax_mag, width=6).pack(side=tk.LEFT, padx=1)

        row6 = ttk.Frame(frame)
        row6.pack(fill=tk.X, pady=2)
        ttk.Label(row6, text="Y-Phase:").pack(side=tk.LEFT, padx=2)
        self.ymin_phase = tk.StringVar(value="-200")
        self.ymax_phase = tk.StringVar(value="200")
        ttk.Entry(row6, textvariable=self.ymin_phase, width=6).pack(side=tk.LEFT, padx=1)
        ttk.Label(row6, text="–").pack(side=tk.LEFT)
        ttk.Entry(row6, textvariable=self.ymax_phase, width=6).pack(side=tk.LEFT, padx=1)

    # ── 绘图控制按钮 ──
    def _build_plot_controls(self, parent):
        frame = ttk.LabelFrame(parent, text="Plot Controls", padding=5)
        frame.pack(fill=tk.X, padx=2, pady=2)

        row1 = ttk.Frame(frame)
        row1.pack(fill=tk.X, pady=2)
        ttk.Button(row1, text="|S| Magnitude", command=lambda: self._on_plot('mag')).pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        ttk.Button(row1, text="Phase", command=lambda: self._on_plot('phase')).pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)

        row2 = ttk.Frame(frame)
        row2.pack(fill=tk.X, pady=2)
        ttk.Button(row2, text="Mag + Phase", command=lambda: self._on_plot('mag_phase')).pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        ttk.Button(row2, text="2×2 Grid", command=lambda: self._on_plot('2x2')).pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)

        row3 = ttk.Frame(frame)
        row3.pack(fill=tk.X, pady=2)
        ttk.Button(row3, text="Save Figure As...", command=self._on_save).pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)

        row4 = ttk.Frame(frame)
        row4.pack(fill=tk.X, pady=2)
        ttk.Button(row4, text="🔬 LC Extract...", command=self._on_lc_extract).pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)

    # ── Matplotlib 画布 ──
    def _build_canvas(self, parent):
        self.canvas_frame = ttk.Frame(parent)
        self.canvas_frame.pack(fill=tk.BOTH, expand=True)

        # 初始空白图
        self.fig = plt.figure(figsize=(6, 4.5), dpi=100)
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.canvas_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # 工具栏
        toolbar = NavigationToolbar2Tk(self.canvas, self.canvas_frame)
        toolbar.update()

    # ── 状态栏 ──
    def _build_statusbar(self):
        self.status_var = tk.StringVar(value="Ready — Add files to begin (CSV-dB, CSV-RI, or TXT paired)")
        statusbar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding=2)
        statusbar.pack(side=tk.BOTTOM, fill=tk.X)

    # ═══════════════════════════════════════════════════
    #  事件处理
    # ═══════════════════════════════════════════════════

    def _on_add_files(self):
        """添加文件"""
        filepaths = filedialog.askopenfilenames(
            title="Select S-Parameter Data Files",
            filetypes=[
                ("All Supported", "*.csv;*.txt;*.dat"),
                ("CSV Files", "*.csv"),
                ("Text Files", "*.txt;*.dat"),
                ("All Files", "*.*"),
            ])
        if not filepaths:
            return

        fmt_mode = self.fmt_var.get()
        count = 0
        for fp in filepaths:
            parsed = load_file(fp, fmt=fmt_mode)
            if parsed is not None:
                name = parsed.get('filename', os.path.basename(fp))
                # 处理重名
                base = name
                idx = 1
                while name in self.loaded_files:
                    name = f"{base} ({idx})"
                    idx += 1
                self.loaded_files[name] = parsed
                self.file_listbox.insert(tk.END, name)
                count += 1
            else:
                messagebox.showwarning("Parse Error",
                                       f"Could not parse file:\n{os.path.basename(fp)}")

        self.status_var.set(f"Loaded {count} file(s). Select a file and click a plot button.")

    def _on_add_folder(self):
        """添加文件夹中所有支持的文件"""
        folder = filedialog.askdirectory(title="Select Folder with Data Files")
        if not folder:
            return

        fmt_mode = self.fmt_var.get()
        count = 0
        for root_dir, _, files in os.walk(folder):
            for f in sorted(files):
                if f.lower().endswith(('.csv', '.txt', '.dat')):
                    fp = os.path.join(root_dir, f)
                    parsed = load_file(fp, fmt=fmt_mode)
                    if parsed is not None:
                        name = parsed.get('filename', f)
                        base = name
                        idx = 1
                        while name in self.loaded_files:
                            name = f"{base} ({idx})"
                            idx += 1
                        self.loaded_files[name] = parsed
                        self.file_listbox.insert(tk.END, name)
                        count += 1

        self.status_var.set(f"Loaded {count} file(s) from folder.")

    def _on_convert_txt_to_lcalpha_csv(self):
        """Choose a CST TXT-export folder and create an RI CSV accepted by LCalpha."""
        # Try several likely default locations
        script_dir = os.path.dirname(os.path.abspath(__file__))
        candidates = [
            os.path.join(script_dir, 'C1', 'S参数导出_20260730'),
            os.path.join(script_dir, 'topology_test_pass', 'S参数导出_20260730'),
            os.path.join(script_dir, 'S_Parameter_Results_pass'),
        ]
        default_folder = next((d for d in candidates if os.path.isdir(d)), os.getcwd())
        folder = filedialog.askdirectory(
            title='Select CST TXT export folder', initialdir=default_folder)
        if not folder:
            return

        output_path = filedialog.asksaveasfilename(
            title='Save LCalpha-compatible CSV', initialdir=folder,
            initialfile='S_parameters_for_LCalpha.csv', defaultextension='.csv',
            filetypes=[('CSV Files', '*.csv')])
        if not output_path:
            return

        try:
            row_count = convert_txt_folder_to_lcalpha_csv(folder, output_path)
        except (OSError, ValueError, TypeError, pd.errors.ParserError) as exc:
            messagebox.showerror('TXT Conversion Failed', str(exc))
            self.status_var.set('TXT to LC CSV conversion failed.')
            return

        self.status_var.set(f'Created LCalpha CSV: {os.path.basename(output_path)} ({row_count} points)')
        messagebox.showinfo(
            'Conversion Complete',
            f'LCalpha-compatible RI CSV created:\n{output_path}\n\n'
            f'{row_count} angle/frequency points; original theta values retained.')

    def _on_remove_file(self):
        """移除选中的文件"""
        selected = self.file_listbox.curselection()
        for idx in reversed(selected):
            name = self.file_listbox.get(idx)
            del self.loaded_files[name]
            self.file_listbox.delete(idx)
        self.status_var.set(f"Removed {len(selected)} file(s).")

    def _on_clear(self):
        """清空所有文件"""
        self.loaded_files.clear()
        self.file_listbox.delete(0, tk.END)
        self.fig.clear()
        self.canvas.draw()
        self.status_var.set("Cleared. Add files to begin.")

    def _get_selected_parsed(self):
        """获取当前选中的文件解析数据（支持多选合并）"""
        selected = self.file_listbox.curselection()
        if not selected:
            # 如果没有选中，使用全部
            if len(self.loaded_files) > 0:
                names = list(self.loaded_files.keys())
            else:
                return None
        else:
            names = [self.file_listbox.get(i) for i in selected]
        return [self.loaded_files[n] for n in names]

    def _get_params(self):
        """收集当前所有绘图参数"""
        pols = []
        if self.pol_te.get():
            pols.append("TE")
        if self.pol_tm.get():
            pols.append("TM")
        pols = tuple(pols) if pols else ("TE",)

        angles_str = self.angles_var.get().strip()
        angles = None
        if angles_str:
            try:
                angles = [float(a.strip()) for a in angles_str.split(',') if a.strip()]
            except ValueError:
                pass

        freq_range = None
        fmin_str = self.fmin_var.get().strip()
        fmax_str = self.fmax_var.get().strip()
        if fmin_str and fmax_str:
            try:
                freq_range = (float(fmin_str), float(fmax_str))
            except ValueError:
                pass

        try:
            ylim_mag = (float(self.ymin_mag.get()), float(self.ymax_mag.get()))
        except ValueError:
            ylim_mag = (-50, 5)

        try:
            ylim_phase = (float(self.ymin_phase.get()), float(self.ymax_phase.get()))
        except ValueError:
            ylim_phase = (-200, 200)

        return {
            's_type': self.stype_var.get(),
            'pols': pols,
            'angles': angles,
            'freq_range': freq_range,
            'ylim_mag': ylim_mag,
            'ylim_phase': ylim_phase,
        }

    def _on_plot(self, plot_type='mag'):
        """执行绘图 — 多文件叠加时只画数据线，装饰（标签/图例/参考线）统一加一次"""
        parsed_list = self._get_selected_parsed()
        if not parsed_list:
            messagebox.showinfo("No Data", "Please add and select data file(s) first.")
            return

        params = self._get_params()
        common = {k: params[k] for k in ('s_type', 'pols', 'angles', 'freq_range')}
        n_files = len(parsed_list)

        # 清空画布
        self.fig.clf()

        if plot_type == 'mag':
            ax = self.fig.add_subplot(111)

            # 逐文件画数据线 (不添加装饰)，每个文件用不同色系
            for idx, parsed in enumerate(parsed_list):
                plot_s_magnitude(parsed, **common, ylim=params['ylim_mag'],
                                 fig=self.fig, ax=ax, decorate=False,
                                 cmap=FILE_CMAPS[idx % len(FILE_CMAPS)])

            # 统一添加装饰（只一次）
            ax.set_xlabel("Frequency (GHz)")
            ax.set_ylabel(f"|{common['s_type']}| (dB)")
            ax.set_ylim(params['ylim_mag'])
            if common['freq_range']:
                ax.set_xlim(common['freq_range'])
            ax.axhline(-3,  color='gray', linestyle=':', linewidth=0.6, alpha=0.5)
            ax.axhline(-10, color='gray', linestyle=':', linewidth=0.6, alpha=0.5)
            ax.grid(True)
            ax.legend(loc='lower left', ncol=2, fontsize=7)
            title = (f"{n_files} files — |{common['s_type']}| (dB)"
                     if n_files > 1 else parsed_list[0].get('filename', ''))
            if title:
                ax.set_title(title, fontsize=9)

        elif plot_type == 'phase':
            ax = self.fig.add_subplot(111)

            for idx, parsed in enumerate(parsed_list):
                plot_s_phase(parsed, **common, ylim=params['ylim_phase'],
                             fig=self.fig, ax=ax, decorate=False,
                             cmap=FILE_CMAPS[idx % len(FILE_CMAPS)])

            ax.set_xlabel("Frequency (GHz)")
            ax.set_ylabel(f"Phase({common['s_type']}) (deg)")
            ax.set_ylim(params['ylim_phase'])
            if common['freq_range']:
                ax.set_xlim(common['freq_range'])
            ax.axhline(0, color='gray', linestyle=':', linewidth=0.6, alpha=0.5)
            ax.grid(True)
            ax.legend(loc='lower left', ncol=2, fontsize=7)
            title = (f"{n_files} files — Phase({common['s_type']}) (deg)"
                     if n_files > 1 else parsed_list[0].get('filename', ''))
            if title:
                ax.set_title(title, fontsize=9)

        elif plot_type == 'mag_phase':
            # 2×1 双面板
            ax_mag = self.fig.add_subplot(2, 1, 1)
            ax_pha = self.fig.add_subplot(2, 1, 2)

            for idx, parsed in enumerate(parsed_list):
                plot_s_mag_and_phase(parsed, **common,
                                     ylim_mag=params['ylim_mag'],
                                     ylim_phase=params['ylim_phase'],
                                     fig=self.fig, axes=(ax_mag, ax_pha),
                                     decorate=False,
                                     cmap=FILE_CMAPS[idx % len(FILE_CMAPS)])

            # 幅值面板装饰
            ax_mag.set_ylabel(f"|{common['s_type']}| (dB)")
            ax_mag.set_ylim(params['ylim_mag'])
            if common['freq_range']:
                ax_mag.set_xlim(common['freq_range'])
            ax_mag.axhline(-3,  color='gray', linestyle=':', linewidth=0.6, alpha=0.5)
            ax_mag.axhline(-10, color='gray', linestyle=':', linewidth=0.6, alpha=0.5)
            ax_mag.grid(True)
            ax_mag.legend(loc='lower left', ncol=3, fontsize=6.5)

            # 相位面板装饰
            ax_pha.set_xlabel("Frequency (GHz)")
            ax_pha.set_ylabel(f"Phase({common['s_type']}) (deg)")
            ax_pha.set_ylim(params['ylim_phase'])
            if common['freq_range']:
                ax_pha.set_xlim(common['freq_range'])
            ax_pha.axhline(0, color='gray', linestyle=':', linewidth=0.6, alpha=0.5)
            ax_pha.grid(True)

            if n_files > 1:
                self.fig.suptitle(f"{n_files} files — {common['s_type']}",
                                  fontsize=9, fontweight='bold', y=0.995)
            else:
                title = parsed_list[0].get('filename', '')
                if title:
                    self.fig.suptitle(title, fontsize=9, fontweight='bold', y=0.995)

        elif plot_type == '2x2':
            # 2×2 四子图 (单个文件，直接用 decorate=True)
            parsed = parsed_list[0]
            title = parsed.get('filename', '')
            axes = self.fig.subplots(2, 2, squeeze=False, figsize=(7.5, 5.2))
            plot_s_2x2(parsed, **common,
                       ylim_mag=params['ylim_mag'],
                       ylim_phase=params['ylim_phase'],
                       title=title, fig=self.fig, axes=axes, decorate=True)
            if n_files > 1:
                self.status_var.set("Note: 2x2 grid shows only the first selected file.")

        self.fig.tight_layout()
        self.canvas.draw()
        self.status_var.set(
            f"Plotted {n_files} file(s) — "
            f"S={common['s_type']}, Pols={list(common['pols'])}, "
            f"Type={plot_type}"
        )

    def _on_save(self):
        """保存当前图形"""
        if self.fig is None:
            messagebox.showinfo("No Figure", "No figure to save.")
            return

        filepath = filedialog.asksaveasfilename(
            title="Save Figure",
            defaultextension=".pdf",
            filetypes=[
                ("PDF", "*.pdf"),
                ("PNG", "*.png"),
                ("SVG", "*.svg"),
                ("JPEG", "*.jpg"),
                ("All Files", "*.*"),
            ])
        if filepath:
            self.fig.savefig(filepath, bbox_inches='tight', dpi=600)
            self.status_var.set(f"Figure saved to: {os.path.basename(filepath)}")

    def _on_lc_extract(self):
        """打开 LC 参数提取窗口"""
        try:
            from lc_extract_window import LCExtractWindow
        except ImportError:
            messagebox.showerror("Import Error",
                "Cannot import lc_extract_window.py.\n"
                "Make sure it is in the same directory as this script.")
            return

        # 获取当前选中文件的路径，传入提取窗口
        selected = self._get_selected_parsed()
        filepaths = []
        if selected:
            for p in selected:
                fp = p.get('filepath', '')
                if fp and fp.lower().endswith('.csv'):
                    filepaths.append(fp)

        win = LCExtractWindow(self.root, filepaths=filepaths)
        win.grab_set()  # 模态窗口

    def _on_about(self):
        messagebox.showinfo(
            "About SDC-ECM",
            "SDC-ECM — Spatial Dispersion Compensated\n"
            "Equivalent Circuit Model Extraction Tool\n\n"
            "S-Parameter Visualization + LC/Alpha Extraction\n"
            "Supports: CST TXT export / CSV (dB / RI)\n"
            "Method: Foster reactance theorem + S21 correction")


# ═══════════════════════════════════════════════════════════
#  入口
# ═══════════════════════════════════════════════════════════
if __name__ == "__main__":
    # 防止 MKL/OpenMP 冲突
    os.environ.setdefault("KMP_DUPLICATE_LIB_OK", "TRUE")
    os.environ.setdefault("OMP_NUM_THREADS", "1")

    root = tk.Tk()
    app = SParameterGUI(root)

    # 自动加载 data 文件夹下的示例数据
    data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
    if os.path.isdir(data_dir):
        cnt = 0
        for fn in sorted(os.listdir(data_dir)):
            fp = os.path.join(data_dir, fn)
            if os.path.isfile(fp) and fn.lower().endswith('.csv'):
                parsed = load_file(fp)
                if parsed is not None:
                    app.loaded_files[fn] = parsed
                    app.file_listbox.insert(tk.END, fn)
                    cnt += 1
            elif os.path.isdir(fp):
                for sub_fn in sorted(os.listdir(fp)):
                    sfp = os.path.join(fp, sub_fn)
                    if os.path.isfile(sfp) and sub_fn.lower().endswith('.txt') and 'Ph' not in sub_fn:
                        parsed = load_file(sfp, fmt='txt_paired')
                        if parsed is not None:
                            name = os.path.join(os.path.basename(fp), sub_fn)
                            app.loaded_files[name] = parsed
                            app.file_listbox.insert(tk.END, name)
                            cnt += 1
        if cnt > 0:
            app.status_var.set(f"Auto-loaded {cnt} example file(s) from data/. Select and plot!")
        else:
            app.status_var.set("Ready — Add files to begin (CSV-dB, CSV-RI, or TXT paired)")
    else:
        app.status_var.set("Ready — Add files to begin (CSV-dB, CSV-RI, or TXT paired)")

    root.mainloop()
