# app.py
# 无形资产分成率模拟器（对数正态分布版） - 含 K-S 拟合优度检验
# 支持均值=4.7%, 标准差=5% | 添加数值稳定性判断 | 4位小数 | Excel导出 | K-S检验

import streamlit as st
import numpy as np
from scipy.stats import lognorm, kstest
import xlsxwriter
from io import BytesIO
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib

# -----------------------------
# 中文支持
# -----------------------------
matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'Arial Unicode MS', 'DejaVu Sans']
matplotlib.rcParams['axes.unicode_minus'] = False

# -----------------------------
# 页面配置
# -----------------------------
st.set_page_config(
    page_title="无形资产分成率模拟器（对数正态+K-S检验）",
    page_icon="📊",
    layout="wide"
)

st.title("📊 无形资产分成率模拟器（对数正态分布 + K-S拟合检验）")

st.markdown("""
> ✅ 支持高波动场景（CV > 1）| ✅ 添加 **K-S 拟合优度检验** | 双输入模式 | 4位小数 | 生成Excel报告
""")

# -----------------------------
# 工具函数
# -----------------------------

def lognormal_params_from_mean_std(mean_percent, std_percent):
    """
    根据均值和标准差计算对数正态分布的 mu 和 sigma
    允许 CV > 1，仅检查数值稳定性
    """
    mean = mean_percent
    std = std_percent

    if std == 0:
        std = 1e-6

    cv_squared = (std / mean) ** 2
    if not np.isfinite(cv_squared) or cv_squared <= 0:
        return None, None

    sigma_squared = np.log(1 + cv_squared)
    if not np.isfinite(sigma_squared) or sigma_squared <= 0:
        return None, None

    sigma = np.sqrt(sigma_squared)
    mu = np.log(mean) - sigma_squared / 2

    if not np.isfinite(mu) or not np.isfinite(sigma):
        return None, None

    return mu, sigma


def lognormal_params_from_median_cv(median_percent, cv):
    """
    根据中位数和变异系数求 mu 和 sigma
    """
    if cv <= 0:
        return None, None

    sigma_squared = np.log(1 + cv**2)
    if not np.isfinite(sigma_squared) or sigma_squared <= 0:
        return None, None

    sigma = np.sqrt(sigma_squared)
    mu = np.log(median_percent)

    if not np.isfinite(mu) or not np.isfinite(sigma):
        return None, None

    return mu, sigma


def compute_detailed_stats(mu, sigma, simulated_rates):
    """
    计算统计量（保留4位小数）
    """
    mean = np.exp(mu + sigma**2 / 2)
    median = np.exp(mu)
    variance = (np.exp(sigma**2) - 1) * np.exp(2*mu + sigma**2)
    std = np.sqrt(variance)
    cv = std / mean if mean != 0 else None

    skew = (np.exp(sigma**2) + 2) * np.sqrt(np.exp(sigma**2) - 1)
    kurt = np.exp(4*sigma**2) + 2*np.exp(3*sigma**2) + 3*np.exp(2*sigma**2) - 6

    mode = np.exp(mu - sigma**2) if sigma < np.sqrt(mu) else None

    p5 = lognorm.ppf(0.05, s=sigma, scale=np.exp(mu))
    p95 = lognorm.ppf(0.95, s=sigma, scale=np.exp(mu))

    return {
        "均值 (%)": round(mean, 4),
        "中位数 (%)": round(median, 4),
        "众数 (%)": round(mode, 4) if mode is not None else "N/A",
        "标准差 (%)": round(std, 4),
        "变异系数 (CV)": round(cv, 4) if cv is not None else "N/A",
        "偏度": round(skew, 4),
        "峰度": round(kurt, 4),
        "P5 (%)": round(p5, 4),
        "P95 (%)": round(p95, 4),
        "模拟样本量": len(simulated_rates)
    }


def perform_ks_test(data, mu, sigma):
    """
    对样本数据执行 K-S 拟合优度检验
    H0: 数据来自指定参数的对数正态分布
    """
    # 理论 CDF 函数（使用估计参数）
    def lognorm_cdf(x):
        return lognorm.cdf(x, s=sigma, scale=np.exp(mu))

    # 执行 K-S 检验
    ks_stat, p_value = kstest(data, lognorm_cdf)

    result = "通过" if p_value > 0.05 else "未通过"
    color = "green" if p_value > 0.05 else "red"

    return {
        "KS 统计量 (D)": round(ks_stat, 6),
        "p-value": round(p_value, 6),
        "检验结果": f"<span style='color:{color}; font-weight:bold;'>{result}</span>"
    }


def create_excel_report(scenario_name, params, data, quantile_dict=None, stats=None, ks_result=None):
    """
    生成 Excel 报告（包含 K-S 检验结果）
    """
    output = BytesIO()
    with xlsxwriter.Workbook(output, {'in_memory': True}) as workbook:
        ws = workbook.add_worksheet("模拟数据")

        title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'fg_color': '#f0f2f6'})
        header_format = workbook.add_format({'bold': True, 'bg_color': '#4e79a7', 'font_color': 'white'})
        float_format = workbook.add_format({'num_format': '0.0000'})

        ws.merge_range('A1:F1', f'【{scenario_name}】对数正态分布模拟报告', title_format)

        params_data = [
            ['参数', '值'],
            ['样本量', params['sample_size']],
            ['mu (μ)', f"{params['mu']:.4f}"],
            ['sigma (σ)', f"{params['sigma']:.4f}"],
            ['中位数 (%)', f"{params['median']:.4f}"],
            ['变异系数 (CV)', f"{params['cv']:.4f}"],
            ['均值 (%)', f"{params['mean']:.4f}"],
            ['标准差 (%)', f"{params['std']:.4f}"]
        ]
        for i, row in enumerate(params_data):
            ws.write(i + 2, 0, row[0])
            ws.write(i + 2, 1, row[1])
        ws.set_column('A:A', 25)
        ws.set_column('B:B', 18)

        if quantile_dict:
            ws.write(10, 0, "分位数分析", header_format)
            ws.write(11, 0, "分位数")
            ws.write(11, 1, "分成率 (%)")
            row_idx = 12
            for label, value in quantile_dict.items():
                ws.write(row_idx, 0, label)
                ws.write(row_idx, 1, value, float_format)
                row_idx += 1

        if stats:
            ws.write(row_idx + 1, 0, "核心统计量", header_format)
            ws.write(row_idx + 2, 0, "统计量")
            ws.write(row_idx + 2, 1, "值")
            stat_row = row_idx + 3
            for key, value in stats.items():
                ws.write(stat_row, 0, key)
                fmt = float_format if isinstance(value, (int, float)) else None
                ws.write(stat_row, 1, value, fmt)
                stat_row += 1

        if ks_result:
            ws.write(stat_row + 1, 0, "K-S 拟合检验", header_format)
            ws.write(stat_row + 2, 0, "指标")
            ws.write(stat_row + 2, 1, "值")
            ws.write(stat_row + 3, 0, "KS 统计量 (D)")
            ws.write(stat_row + 3, 1, ks_result["KS 统计量 (D)"])
            ws.write(stat_row + 4, 0, "p-value")
            ws.write(stat_row + 4, 1, ks_result["p-value"])
            ws.write(stat_row + 5, 0, "结论")
            ws.write(stat_row + 5, 1, ks_result["检验结果"].replace('<span style=', '').replace('</span>', ''))

        ws.write(2, 3, '模拟分成率 (%)')
        for i, rate in enumerate(data, start=3):
            ws.write(i, 3, rate, float_format)

        chart = workbook.add_chart({'type': 'line'})
        chart.add_series({
            'values': f'=模拟数据!$D$3:$D${len(data)+2}',
            'name': '分成率模拟数据'
        })
        chart.set_title({'name': '对数正态分布模拟结果'})
        chart.set_x_axis({'name': '样本序号'})
        chart.set_y_axis({'name': '分成率 (%)'})
        ws.insert_chart('F1', chart)

    output.seek(0)
    return output


# -----------------------------
# 用户输入（侧边栏）
# -----------------------------
st.sidebar.header("🔧 模拟参数设置")

sample_size = st.sidebar.number_input(
    "样本量",
    min_value=10,
    max_value=1000,
    value=136,
    step=1
)

modeling_basis = st.sidebar.radio(
    "建模基准",
    options=["均值 + 标准差", "中位数 + 变异系数"],
    help="对数正态分布适合右偏数据"
)

mu, sigma = None, None

if modeling_basis == "均值 + 标准差":
    mean_rate = st.sidebar.number_input(
        "平均分成率 (%)", min_value=0.0001, max_value=50.0, value=4.7, step=0.1, format="%.4f"
    )
    std_rate = st.sidebar.number_input(
        "标准差 (%)", min_value=0.0001, max_value=30.0, value=5.0, step=0.1, format="%.4f"
    )
    mu, sigma = lognormal_params_from_mean_std(mean_rate, std_rate)
    median_rate = np.exp(mu) if mu is not None else None
else:
    median_rate = st.sidebar.number_input(
        "中位数分成率 (%)", min_value=0.0001, max_value=50.0, value=6.0, step=0.1, format="%.4f"
    )
    cv_value = st.sidebar.number_input(
        "变异系数 (CV)", min_value=0.0001, max_value=5.0, value=0.75, step=0.05, format="%.4f"
    )
    mu, sigma = lognormal_params_from_median_cv(median_rate, cv_value)
    mean_rate = np.exp(mu + sigma**2 / 2) if mu is not None else None
    std_rate = np.sqrt((np.exp(sigma**2) - 1) * np.exp(2*mu + sigma**2)) if mu is not None else None

# 参数检查
if mu is None or sigma is None:
    st.error("❌ 无法构建对数正态分布：输入参数导致数值不稳定。")
    st.info("💡 建议：检查输入是否合理，或尝试使用「中位数 + 变异系数」模式。")
    st.stop()

# 显示分布参数
with st.sidebar:
    st.markdown("---")
    st.markdown("**📊 对数正态分布参数**")
    st.code(f"μ = {mu:.4f}\nσ = {sigma:.4f}")

# -----------------------------
# 自定义分位数
# -----------------------------
st.sidebar.markdown("---")
st.sidebar.subheader("🎯 自定义分位数")
quantile_input = st.sidebar.text_input(
    "输入分位数（百分比，用逗号分隔）",
    value="5, 10, 25, 50, 75, 90, 95",
    help="例如：5,10,25,50,75,90,95"
)

try:
    user_quantiles = [float(x.strip()) for x in quantile_input.split(",")]
    user_quantiles = [q for q in user_quantiles if 0 < q < 100]
    if len(user_quantiles) == 0:
        user_quantiles = [5, 10, 25, 50, 75, 90, 95]
except:
    user_quantiles = [5, 10, 25, 50, 75, 90, 95]

# -----------------------------
# 主界面：绘图与分析
# -----------------------------
simulated_rates = lognorm.rvs(s=sigma, scale=np.exp(mu), size=sample_size)
stats = compute_detailed_stats(mu, sigma, simulated_rates)

# 执行 K-S 检验
ks_result = perform_ks_test(simulated_rates, mu, sigma)

x = np.linspace(0.1, 30, 1000)
pdf = lognorm.pdf(x, s=sigma, scale=np.exp(mu))
cdf = lognorm.cdf(x, s=sigma, scale=np.exp(mu))

col1, col2 = st.columns(2)

with col1:
    st.subheader("概率密度函数 (PDF)")
    fig1, ax1 = plt.subplots(figsize=(6, 4))
    ax1.plot(x, pdf, color='blue', linewidth=2, label='PDF')
    ax1.axvline(mean_rate, color='red', linestyle='--', alpha=0.8, label=f'均值 ({mean_rate:.4f}%)')
    ax1.axvline(median_rate, color='green', linestyle='-.', alpha=0.8, label=f'中位数 ({median_rate:.4f}%)')
    if stats["众数 (%)"] != "N/A":
        ax1.axvline(stats["众数 (%)"], color='orange', linestyle=':', alpha=0.8, label=f'众数 ({stats["众数 (%)"]}%)')
    ax1.axvspan(stats["P5 (%)"], stats["P95 (%)"], color='gray', alpha=0.2, label='95% 区间')
    ax1.set_xlabel('分成率 (%)')
    ax1.set_ylabel('概率密度')
    ax1.set_title('PDF')
    ax1.grid(True, alpha=0.3)
    ax1.legend(fontsize=7)
    st.pyplot(fig1)

with col2:
    st.subheader("累积分布函数 (CDF)")
    fig2, ax2 = plt.subplots(figsize=(6, 4))
    ax2.plot(x, cdf, color='green', linewidth=2, label='CDF')
    ax2.axhline(0.5, color='gray', linestyle='--', alpha=0.5, label='中位数')
    for q in [5, 95]:
        ax2.axhline(q/100, color='gray', linestyle=':', alpha=0.5)
    ax2.set_xlabel('分成率 (%)')
    ax2.set_ylabel('累积概率')
    ax2.set_title('CDF')
    ax2.grid(True, alpha=0.3)
    ax2.legend()
    st.pyplot(fig2)

# -----------------------------
# 统计量与警告
# -----------------------------
cv = stats["变异系数 (CV)"]
if isinstance(cv, (int, float)):
    if cv > 1:
        st.warning(f"⚠️ 警告：变异系数 CV = {cv:.4f} > 1，表示波动性极高，结果仅供参考，请谨慎使用。")
    elif cv > 0.5:
        st.info(f"ℹ️ 变异系数 CV = {cv:.4f}，波动性较高。")

st.markdown("### 📊 核心统计量分析")
stats_df = pd.DataFrame(list(stats.items()), columns=["统计量", "值"])
st.table(stats_df)

# -----------------------------
# 分位数分析
# -----------------------------
st.markdown("### 📊 分位数分析")
quantile_probs = [q / 100 for q in sorted(user_quantiles)]
quantile_values = lognorm.ppf(quantile_probs, s=sigma, scale=np.exp(mu))
quantile_values_rounded = [round(v, 4) for v in quantile_values]
quantile_labels = [f"P{int(q)}" for q in sorted(user_quantiles)]
quantile_dict = dict(zip(quantile_labels, quantile_values_rounded))

quantile_df = pd.DataFrame({
    "分位数": quantile_labels,
    "分成率 (%)": [f"{v:.4f}" for v in quantile_values_rounded]
})
st.table(quantile_df)

# -----------------------------
# K-S 拟合优度检验
# -----------------------------
st.markdown("### 🧪 K-S 拟合优度检验")
ks_df = pd.DataFrame([ks_result])
# 使用 unsafe_allow_html 渲染带颜色的结果
st.markdown(ks_df.to_html(escape=False, index=False), unsafe_allow_html=True)

st.markdown("""
- **H₀**: 样本数据来自指定参数的对数正态分布  
- **H₁**: 样本数据不来自该分布  
- **判断标准**: p > 0.05 → 接受 H₀（拟合良好）
""")

# -----------------------------
# Excel 报告
# -----------------------------
st.markdown("---")
st.subheader("📥 生成 Excel 报告")

if st.button("✨ 生成并下载 Excel 报告"):
    excel_data = create_excel_report(
        "对数正态分布模拟",
        {
            'sample_size': sample_size,
            'mu': mu,
            'sigma': sigma,
            'median': median_rate,
            'cv': cv_value if modeling_basis == "中位数 + 变异系数" else std_rate / mean_rate,
            'mean': mean_rate,
            'std': std_rate
        },
        simulated_rates,
        quantile_dict=quantile_dict,
        stats=stats,
        ks_result=ks_result
    )
    st.download_button(
        label="📥 点击下载 Excel 文件",
        data=excel_data,
        file_name=f"对数正态_分成率模拟_{mean_rate:.4f}%.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.success("✅ Excel 报告已生成！点击上方按钮下载。")

# -----------------------------
# 尾注
# -----------------------------
st.markdown("<br><hr><center>Powered by Streamlit | 开发者：Qwen | 含 K-S 拟合检验</center>", unsafe_allow_html=True)