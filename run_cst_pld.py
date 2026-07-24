#pip install pywin32 numpy pandas
import os
import math
import numpy as np
import pandas as pdd
import win32com.client
import time
import shutil
from Modeling_for_TO_old import *

def delete_directory_robustly(dir_path):
    """强健的目录删除函数，应对 Windows 文件锁死延迟"""
    if not os.path.exists(dir_path):
        return

    for attempt in range(5):
        try:
            shutil.rmtree(dir_path)
            return
        except Exception as e:
            if attempt < 4:
                print(f"    [重试 {attempt + 1}/5] 删除目录失败: {e}，等待 1 秒后重试...")
                time.sleep(1)
            else:
                print(f"    [警告] 无法彻底删除临时工作目录: {dir_path}")


def merge_sparameter_data(existing_path, new_data_df):
    """将新提取的0°和80° S参数数据合并到已有CSV文件中"""
    if os.path.exists(existing_path):
        existing_df = pdd.read_csv(existing_path)
        # 移除已有的0°和80°数据（避免重复）
        existing_df = existing_df[~existing_df['Angle'].isin([0, 80])]
        print(f"    >> 已有数据: {len(existing_df)} 行 (角度: {sorted(existing_df['Angle'].unique())})")
        # 合并新旧数据
        merged_df = pdd.concat([existing_df, new_data_df], ignore_index=True)
    else:
        print(f"    >> 目标文件不存在，将创建新文件: {existing_path}")
        merged_df = new_data_df

    # 按角度和频率排序
    merged_df.sort_values(by=['Angle', 'Frequency_GHz'], inplace=True)
    # 统一列顺序
    cols_order = ['Angle', 'Frequency_GHz',
                  'S11_TE_Mag_dB', 'S11_TE_Phase_deg', 'S21_TE_Mag_dB', 'S21_TE_Phase_deg',
                  'S11_TM_Mag_dB', 'S11_TM_Phase_deg', 'S21_TM_Mag_dB', 'S21_TM_Phase_deg']
    valid_cols = [c for c in cols_order if c in merged_df.columns]
    merged_df = merged_df[valid_cols]
    merged_df.to_csv(existing_path, index=False)
    print(f"    >> 合并后: {len(merged_df)} 行 (角度: {sorted(merged_df['Angle'].unique())})")
    return merged_df


# ==============================================================================
# 3. 主程序调度流
# ==============================================================================
def run_sweep_process():
    """主程序：自动化仿真流程"""
    csv_file = "cst_sweep_data.csv"
    if not os.path.exists(csv_file):
        print(f"❌ 找不到配置文件: {csv_file}")
        return

    df = pdd.read_csv(csv_file)
    print(f"✅ 加载待扫参数据: {len(df)} 条")

    base_dir = os.getcwd()#获取当前文件夹路径
    res_dir = os.path.join(base_dir, "S_Parameter_Results")
    temp_prj_dir = os.path.join(base_dir, "Temp_Projects")

    if not os.path.exists(res_dir):
        os.makedirs(res_dir)
    if not os.path.exists(temp_prj_dir):
        os.makedirs(temp_prj_dir)


    total = len(df)
    print(f"\n🚀 开始执行自动化仿真 (共 {total} 条)...\n")

    cst = None
    builder = None
    sweep = None
    extractor = None

    try:
        # 初始化 CST 实例
        cst = CSTInterface(label="New", project_name="topology_test")

        builder = SimulationBuilder(cst)
        sweep = SweepManager(cst)
        extractor = ResultExtractor(cst)

        for count, (_, row) in enumerate(df.iterrows()):

            r_val = round(float(row['Var_Radius_mm']), 2)
            w_val = round(float(row['Var_Width_mm']), 2)
            sym_type = int(row['Sym_Type'])
            source_id = str(row['Source_ID']).split('.')[0]

            # 输出文件名: ID{Source_ID}_{Sym_Type}_R{R}_W{W}.csv
            out_name = f"ID{source_id}_{sym_type}_R{r_val}_W{w_val}.csv"

            print(f"\n[{count + 1}/{total}] {out_name}")

            # 跳过已存在的文件（断点续跑）
            out_path = os.path.join(res_dir, out_name)
            if os.path.exists(out_path):
                print(f"    -> 已存在，跳过")
                continue

            # 获取路径数据
            path_cols = [c for c in df.columns if c.startswith('Path_')]
            path_cols.sort(key=lambda x: int(x.split('_')[1]))
            raw_vals = row[path_cols].values.astype(float)
            points = raw_vals[~np.isnan(raw_vals)].reshape(-1, 2)

            # 构建几何和配置
            print(f"    -> 初始化环境与建模...")
            is_build_success = builder.build(row, points)

            if not is_build_success:
                print(f"    ⚠️ [触发拦截] 拓扑 {source_id} 实体建模失败，跳过。")
                continue

            # 运行参数扫描
            print(f"    -> 运行仿真扫描...")
            is_sim_success = sweep.run_sweep_task()

            # 提取结果并合并到已有CSV
            if is_sim_success:
                try:
                    print(f"    -> 提取仿真结果 (0°和80°)...")
                    temp_name = f"_temp_{out_name}"
                    extractor.extract_db_results(res_dir, temp_name)

                    temp_path = os.path.join(res_dir, temp_name)
                    if os.path.exists(temp_path):
                        new_df = pdd.read_csv(temp_path)
                        merge_sparameter_data(out_path, new_df)
                        os.remove(temp_path)
                        print(f"    ✅ 已保存到 {out_name}")
                    else:
                        print(f"    ❌ 临时提取文件未生成: {temp_path}")
                except Exception as e:
                    print(f"    ❌ 提取/合并结果失败: {e}")

            time.sleep(1)

        print(f"\n🎉 所有任务执行完毕！")
        print(f"📊 数据保存在: {res_dir}")
        print(f"💡 临时工程保存在: {temp_prj_dir}")

    except Exception as e:
        print(f"\n❌ 发生致命错误: {e}")
    finally:
        # 确保 CST 正确关闭
        if cst is not None:
            try:
                print("\n正在关闭 CST...")
                cst.quit_project_without_saving()
                print("✓ CST 已关闭")
            except Exception as e:
                print(f"[警告] 关闭 CST 时出现问题: {e}")


if __name__ == "__main__":
    run_sweep_process()
