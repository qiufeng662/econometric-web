# -*- coding: utf-8 -*-
"""
计量经济学方法模块
支持各种DID方法、IV、PSM、RDD、SCM等
"""

import pandas as pd
import numpy as np
import statsmodels.api as sm
from statsmodels.formula.api import ols
from statsmodels.discrete.discrete_model import Logit
from statsmodels.stats.outliers_influence import variance_inflation_factor
import warnings
from datetime import datetime
import os

warnings.filterwarnings('ignore')

# ============================================================
# DID 方法
# ============================================================

def classic_did(df, dep_var, treat_var, time_var, post_period, id_var=None, controls=None):
    """
    经典DID方法
    
    Parameters:
    -----------
    df : DataFrame
        数据
    dep_var : str
        因变量
    treat_var : str
        处理变量（0/1）
    time_var : str
        时间变量
    post_period : int/float
        处理期开始的时间点
    id_var : str
        个体标识变量（可选）
    controls : list
        控制变量列表
    
    Returns:
    --------
    dict : DID结果
    """
    df_did = df.copy()
    
    # 创建时间哑变量
    df_did['post'] = (df_did[time_var] >= post_period).astype(int)
    
    # 创建处理组哑变量
    df_did['treated'] = df_did[treat_var]
    
    # 创建交互项
    df_did['did'] = df_did['treated'] * df_did['post']
    
    # 构建公式
    formula = f"{dep_var} ~ treated + post + did"
    if controls:
        formula += " + " + " + ".join(controls)
    
    # 回归
    model = ols(formula=formula, data=df_did).fit(cov_type='HC3')
    
    # 平行趋势检验（处理前时期）
    pre_trend = df_did[df_did[time_var] < post_period]
    if len(pre_trend) > 0:
        pre_model = ols(formula=formula, data=pre_trend).fit(cov_type='HC3')
        pre_trend_p = pre_model.pvalues.get('did', 1.0)
    else:
        pre_trend_p = None
    
    return {
        'model': model,
        'did_coef': model.params.get('did', np.nan),
        'did_se': model.bse.get('did', np.nan),
        'did_pvalue': model.pvalues.get('did', np.nan),
        'pre_trend_p': pre_trend_p,
        'n_obs': len(df_did),
        'method': '经典DID'
    }

def multiple_period_did(df, dep_var, treat_var, time_var, treat_time_var, controls=None):
    """
    多期DID方法（允许不同个体在不同时间接受处理）
    
    Parameters:
    -----------
    df : DataFrame
        数据
    dep_var : str
        因变量
    treat_var : str
        处理变量（0/1）
    time_var : str
        时间变量
    treat_time_var : str
        接受处理的时间变量
    controls : list
        控制变量列表
    
    Returns:
    --------
    dict : 多期DID结果
    """
    df_did = df.copy()
    
    # 创建相对时间变量
    df_did['relative_time'] = df_did[time_var] - df_did[treat_time_var]
    
    # 创建处理后哑变量
    df_did['post'] = (df_did['relative_time'] >= 0).astype(int)
    
    # 构建公式（使用相对时间作为因子变量）
    formula = f"{dep_var} ~ C(relative_time) + {treat_var}:C(relative_time)"
    if controls:
        formula += " + " + " + ".join(controls)
    
    # 回归
    model = ols(formula=formula, data=df_did).fit(cov_type='HC3')
    
    return {
        'model': model,
        'n_obs': len(df_did),
        'method': '多期DID'
    }

def event_study_did(df, dep_var, treat_var, time_var, treat_time_var, 
                   time_window=(-5, 5), controls=None):
    """
    事件研究DID方法
    
    Parameters:
    -----------
    df : DataFrame
        数据
    dep_var : str
        因变量
    treat_var : str
        处理变量（0/1）
    time_var : str
        时间变量
    treat_time_var : str
        接受处理的时间变量
    time_window : tuple
        时间窗口 (min, max)
    controls : list
        控制变量列表
    
    Returns:
    --------
    dict : 事件研究结果
    """
    df_event = df.copy()
    
    # 创建相对时间
    df_event['event_time'] = df_event[time_var] - df_event[treat_time_var]
    
    # 限制时间窗口
    df_event = df_event[(df_event['event_time'] >= time_window[0]) & 
                       (df_event['event_time'] <= time_window[1])]
    
    # 创建相对时间哑变量（以-1期为基准）
    df_event['event_time_factor'] = pd.Categorical(df_event['event_time'], 
                                                  categories=sorted(df_event['event_time'].unique()))
    
    # 构建公式
    formula = f"{dep_var} ~ C(event_time_factor)"
    if controls:
        formula += " + " + " + ".join(controls)
    
    # 回归
    model = ols(formula=formula, data=df_event).fit(cov_type='HC3')
    
    # 提取各期系数
    coefs = {}
    ses = {}
    pvalues = {}
    
    for time_point in sorted(df_event['event_time'].unique()):
        if time_point != -1:  # -1期作为基准
            param_name = f"C(event_time_factor)[T.{time_point}]"
            coefs[time_point] = model.params.get(param_name, np.nan)
            ses[time_point] = model.bse.get(param_name, np.nan)
            pvalues[time_point] = model.pvalues.get(param_name, np.nan)
    
    return {
        'model': model,
        'coefs': coefs,
        'ses': ses,
        'pvalues': pvalues,
        'time_points': sorted(df_event['event_time'].unique()),
        'n_obs': len(df_event),
        'method': '事件研究DID'
    }

def difference_in_difference_in_differences(df, dep_var, treat_var1, treat_var2, 
                                         time_var, post_period, controls=None):
    """
    差异中的差异（DDD）方法
    
    Parameters:
    -----------
    df : DataFrame
        数据
    dep_var : str
        因变量
    treat_var1 : str
        第一维处理变量
    treat_var2 : str
        第二维处理变量
    time_var : str
        时间变量
    post_period : int/float
        处理期开始时间
    controls : list
        控制变量列表
    
    Returns:
    --------
    dict : DDD结果
    """
    df_ddd = df.copy()
    
    # 创建时间哑变量
    df_ddd['post'] = (df_ddd[time_var] >= post_period).astype(int)
    
    # 创建三重交互项
    df_ddd['ddd'] = df_ddd[treat_var1] * df_ddd[treat_var2] * df_ddd['post']
    
    # 构建公式
    formula = f"{dep_var} ~ {treat_var1} + {treat_var2} + post + {treat_var1}:{treat_var2} + {treat_var1}:post + {treat_var2}:post + ddd"
    if controls:
        formula += " + " + " + ".join(controls)
    
    # 回归
    model = ols(formula=formula, data=df_ddd).fit(cov_type='HC3')
    
    return {
        'model': model,
        'ddd_coef': model.params.get('ddd', np.nan),
        'ddd_se': model.bse.get('ddd', np.nan),
        'ddd_pvalue': model.pvalues.get('ddd', np.nan),
        'n_obs': len(df_ddd),
        'method': 'DDD'
    }

# ============================================================
# 工具变量法 (IV)
# ============================================================

def instrumental_variables(df, dep_var, endog_var, instruments, controls=None):
    """
    工具变量法（2SLS）
    
    Parameters:
    -----------
    df : DataFrame
        数据
    dep_var : str
        因变量
    endog_var : str
        内生变量
    instruments : list
        工具变量列表
    controls : list
        外生控制变量列表
    
    Returns:
    --------
    dict : IV结果
    """
    # 第一阶段：内生变量对工具变量和控制变量回归
    first_stage_vars = instruments + (controls if controls else [])
    first_formula = f"{endog_var} ~ " + " + ".join(first_stage_vars)
    
    first_model = ols(formula=first_formula, data=df).fit()
    
    # 计算F统计量（弱工具变量检验）
    f_stat = first_model.fvalue
    f_pvalue = first_model.f_pvalue
    
    # 获取拟合值
    df_iv = df.copy()
    df_iv[f'{endog_var}_fitted'] = first_model.fittedvalues
    
    # 第二阶段：因变量对拟合的内生变量和控制变量回归
    second_formula = f"{dep_var} ~ {endog_var}_fitted"
    if controls:
        second_formula += " + " + " + ".join(controls)
    
    second_model = ols(formula=second_formula, data=df_iv).fit(cov_type='HC3')
    
    # IV估计量
    iv_coef = second_model.params[f'{endog_var}_fitted']
    iv_se = second_model.bse[f'{endog_var}_fitted']
    iv_pvalue = second_model.pvalues[f'{endog_var}_fitted']
    
    # 过度识别检验（如果工具变量多于内生变量）
    if len(instruments) > 1:
        # 这里简化处理，实际应该用更复杂的检验
        sargan_p = None
    else:
        sargan_p = None
    
    return {
        'first_stage': first_model,
        'second_stage': second_model,
        'iv_coef': iv_coef,
        'iv_se': iv_se,
        'iv_pvalue': iv_pvalue,
        'f_stat': f_stat,
        'f_pvalue': f_pvalue,
        'sargan_p': sargan_p,
        'n_obs': len(df),
        'method': '工具变量法'
    }

# ============================================================
# 倾向得分匹配 (PSM)
# ============================================================

def propensity_score_matching(df, treat_var, covariates, caliper=0.05, method='nearest'):
    """
    倾向得分匹配
    
    Parameters:
    -----------
    df : DataFrame
        数据
    treat_var : str
        处理变量
    covariates : list
        用于匹配的协变量
    caliper : float
        匹配容忍度
    method : str
        匹配方法 ('nearest', 'radius')
    
    Returns:
    --------
    dict : PSM结果
    """
    df_psm = df.copy()
    
    # 1. 估计倾向得分
    logit_formula = f"{treat_var} ~ " + " + ".join(covariates)
    ps_model = Logit.from_formula(logit_formula, data=df_psm).fit(disp=False)
    df_psm['pscore'] = ps_model.predict()
    
    # 2. 匹配
    treated = df_psm[df_psm[treat_var] == 1]
    control = df_psm[df_psm[treat_var] == 0]
    
    matched_control_indices = []
    
    for idx, treated_row in treated.iterrows():
        pscore_diff = abs(control['pscore'] - treated_row['pscore'])
        
        if method == 'nearest':
            # 最近邻匹配
            min_idx = pscore_diff.idxmin()
            if pscore_diff[min_idx] <= caliper:
                matched_control_indices.append(min_idx)
        elif method == 'radius':
            # 半径匹配
            candidates = control[pscore_diff <= caliper]
            if len(candidates) > 0:
                matched_control_indices.extend(candidates.index.tolist())
    
    # 3. 获取匹配后的数据
    matched_control = control.loc[matched_control_indices].drop_duplicates()
    matched_data = pd.concat([treated, matched_control])
    
    # 4. 平衡性检验
    balance_stats = {}
    for cov in covariates:
        treated_mean = treated[cov].mean()
        control_mean = matched_control[cov].mean()
        std_diff = (treated_mean - control_mean) / treated[cov].std()
        balance_stats[cov] = {
            'treated_mean': treated_mean,
            'control_mean': control_mean,
            'std_diff': std_diff
        }
    
    return {
        'ps_model': ps_model,
        'matched_data': matched_data,
        'balance_stats': balance_stats,
        'n_treated': len(treated),
        'n_matched_control': len(matched_control),
        'method': 'PSM'
    }

def psm_att_estimation(matched_data, dep_var, treat_var, covariates):
    """
    基于匹配数据的ATT估计
    
    Parameters:
    -----------
    matched_data : DataFrame
        匹配后的数据
    dep_var : str
        因变量
    treat_var : str
        处理变量
    covariates : list
        协变量
    
    Returns:
    --------
    dict : ATT估计结果
    """
    # 简单均值比较
    treated_mean = matched_data[matched_data[treat_var] == 1][dep_var].mean()
    control_mean = matched_data[matched_data[treat_var] == 0][dep_var].mean()
    att = treated_mean - control_mean
    
    # 回归调整
    reg_formula = f"{dep_var} ~ {treat_var} + " + " + ".join(covariates)
    reg_model = ols(formula=reg_formula, data=matched_data).fit(cov_type='HC3')
    
    att_reg = reg_model.params[treat_var]
    att_se = reg_model.bse[treat_var]
    att_pvalue = reg_model.pvalues[treat_var]
    
    return {
        'simple_att': att,
        'reg_att': att_reg,
        'reg_se': att_se,
        'reg_pvalue': att_pvalue,
        'n_obs': len(matched_data),
        'method': 'PSM_ATT'
    }

# ============================================================
# 断点回归 (RDD)
# ============================================================

def regression_discontinuity_design(df, dep_var, running_var, cutoff, 
                                  bandwidth=None, controls=None):
    """
    断点回归设计
    
    Parameters:
    -----------
    df : DataFrame
        数据
    dep_var : str
        因变量
    running_var : str
        运行变量
    cutoff : float
        断点值
    bandwidth : float
        带宽（可选）
    controls : list
        控制变量
    
    Returns:
    --------
    dict : RDD结果
    """
    df_rdd = df.copy()
    
    # 创建处理变量
    df_rdd['treated'] = (df_rdd[running_var] >= cutoff).astype(int)
    
    # 创建断点附近的样本（如果指定带宽）
    if bandwidth:
        df_rdd = df_rdd[abs(df_rdd[running_var] - cutoff) <= bandwidth]
    
    # 多项式项
    df_rdd['running_centered'] = df_rdd[running_var] - cutoff
    df_rdd['running_sq'] = df_rdd['running_centered'] ** 2
    
    # 构建公式
    formula = f"{dep_var} ~ treated + running_centered + running_sq + treated:running_centered + treated:running_sq"
    if controls:
        formula += " + " + " + ".join(controls)
    
    # 回归
    model = ols(formula=formula, data=df_rdd).fit(cov_type='HC3')
    
    # LATE估计量
    late = model.params.get('treated', np.nan)
    
    return {
        'model': model,
        'late': late,
        'late_se': model.bse.get('treated', np.nan),
        'late_pvalue': model.pvalues.get('treated', np.nan),
        'n_obs': len(df_rdd),
        'bandwidth': bandwidth,
        'cutoff': cutoff,
        'method': 'RDD'
    }

# ============================================================
# 合成控制法 (SCM)
# ============================================================

def synthetic_control_method(df, dep_var, treat_unit, time_var, treat_time,
                            control_units=None, predictors=None):
    """
    合成控制法（简化版本）
    
    Parameters:
    -----------
    df : DataFrame
        数据（长格式面板数据）
    dep_var : str
        因变量
    treat_unit : str/int
        处理单元标识
    time_var : str
        时间变量
    treat_time : int/float
        处理开始时间
    control_units : list
        对照单元列表
    predictors : list
        预测变量列表
    
    Returns:
    --------
    dict : SCM结果
    """
    # 这里提供简化版本，实际SCM需要更复杂的优化算法
    # 完整实现需要使用专门的SCM包
    
    df_scm = df.copy()
    
    # 筛选前处理期数据
    pre_data = df_scm[df_scm[time_var] < treat_time]
    
    if control_units is None:
        control_units = df_scm[df_scm[treat_unit] != treat_unit][treat_unit].unique()
    
    # 简单合成：使用均值作为合成控制
    control_data = pre_data[pre_data[treat_unit].isin(control_units)]
    
    if predictors:
        # 使用预测变量进行加权
        weights = {}
        treated_pre = pre_data[pre_data[treat_unit] == treat_unit]
        
        for pred in predictors:
            treated_val = treated_pre[pred].mean()
            control_vals = control_data.groupby(treat_unit)[pred].mean()
            
            # 计算权重（距离倒数）
            distances = abs(control_vals - treated_val)
            weights[pred] = 1 / (distances + 0.001)  # 避免除零
            weights[pred] = weights[pred] / weights[pred].sum()
        
        # 平均权重
        avg_weights = pd.Series(index=control_units, dtype=float)
        for unit in control_units:
            unit_weights = [weights[pred][unit] for pred in predictors]
            avg_weights[unit] = np.mean(unit_weights)
        
        avg_weights = avg_weights / avg_weights.sum()
    else:
        # 简单平均
        avg_weights = pd.Series(1/len(control_units), index=control_units)
    
    # 构建合成控制序列
    synthetic_outcome = {}
    for t in df_scm[time_var].unique():
        time_data = df_scm[df_scm[time_var] == t]
        control_outcomes = time_data[time_data[treat_unit].isin(control_units)].set_index(treat_unit)[dep_var]
        synthetic_outcome[t] = (control_outcomes * avg_weights).sum()
    
    # 计算干预效应
    treated_outcome = df_scm[df_scm[treat_unit] == treat_unit].set_index(time_var)[dep_var]
    
    effects = {}
    for t in treated_outcome.index:
        if t >= treat_time:
            effects[t] = treated_outcome[t] - synthetic_outcome.get(t, np.nan)
    
    return {
        'synthetic_outcome': synthetic_outcome,
        'treated_outcome': treated_outcome.to_dict(),
        'effects': effects,
        'weights': avg_weights.to_dict(),
        'n_control_units': len(control_units),
        'method': 'SCM'
    }

# ============================================================
# 其他计量方法
# ============================================================

def quantile_regression(df, dep_var, treat_var, controls=None, quantiles=[0.25, 0.5, 0.75]):
    """
    分位数回归
    
    Parameters:
    -----------
    df : DataFrame
        数据
    dep_var : str
        因变量
    treat_var : str
        核心自变量
    controls : list
        控制变量
    quantiles : list
        分位数列表
    
    Returns:
    --------
    dict : 分位数回归结果
    """
    try:
        import statsmodels.formula.api as smf
        from statsmodels.regression.quantile_regression import QuantReg
    except ImportError:
        return {'error': '需要安装statsmodels支持分位数回归'}
    
    results = {}
    
    for q in quantiles:
        formula = f"{dep_var} ~ {treat_var}"
        if controls:
            formula += " + " + " + ".join(controls)
        
        model = smf.quantreg(formula, data=df).fit(q=q)
        
        results[f'q{q}'] = {
            'coef': model.params[treat_var],
            'se': model.bse[treat_var],
            'pvalue': model.pvalues[treat_var],
            'pseudo_r2': model.prsquared
        }
    
    return {
        'results': results,
        'quantiles': quantiles,
        'n_obs': len(df),
        'method': '分位数回归'
    }

def panel_data_models(df, dep_var, treat_var, controls, id_var, time_var, 
                     model_type='fixed_effects'):
    """
    面板数据模型
    
    Parameters:
    -----------
    df : DataFrame
        面板数据
    dep_var : str
        因变量
    treat_var : str
        自变量
    controls : list
        控制变量
    id_var : str
        个体标识变量
    time_var : str
        时间标识变量
    model_type : str
        模型类型 ('fixed_effects', 'random_effects', 'between', 'within')
    
    Returns:
    --------
    dict : 面板数据模型结果
    """
    try:
        import linearmodels as lm
    except ImportError:
        return {'error': '需要安装linearmodels包'}
    
    # 设置多重索引
    df_panel = df.set_index([id_var, time_var])
    
    # 构建公式
    formula = f"{dep_var} ~ {treat_var}"
    if controls:
        formula += " + " + " + ".join(controls)
    
    if model_type == 'fixed_effects':
        model = lm.PanelOLS.from_formula(formula, data=df_panel)
    elif model_type == 'random_effects':
        model = lm.RandomEffects.from_formula(formula, data=df_panel)
    elif model_type == 'between':
        model = lm.BetweenOLS.from_formula(formula, data=df_panel)
    elif model_type == 'within':
        model = lm.PooledOLS.from_formula(formula, data=df_panel)
    else:
        return {'error': '不支持的面板数据模型类型'}
    
    result = model.fit(cov_type='robust')
    
    return {
        'model': result,
        'coef': result.params.get(treat_var, np.nan),
        'se': result.std_errors.get(treat_var, np.nan),
        'pvalue': result.pvalues.get(treat_var, np.nan),
        'r2': result.rsquared,
        'n_obs': result.nobs,
        'method': f'面板数据({model_type})'
    }

# ============================================================
# 结果格式化
# ============================================================

def format_econometric_results(results_dict):
    """
    格式化计量经济学结果为显示格式
    
    Parameters:
    -----------
    results_dict : dict
        结果字典
    
    Returns:
    --------
    dict : 格式化的结果
    """
    formatted = {}
    
    for method, result in results_dict.items():
        if method == 'classic_did':
            formatted['经典DID'] = {
                'DID系数': f"{result['did_coef']:.4f}",
                '标准误': f"({result['did_se']:.4f})",
                'P值': f"{result['did_pvalue']:.4f}",
                '显著性': '***' if result['did_pvalue'] < 0.01 else '**' if result['did_pvalue'] < 0.05 else '*' if result['did_pvalue'] < 0.1 else '',
                '样本量': result['n_obs'],
                '平行趋势检验P值': f"{result['pre_trend_p']:.4f}" if result['pre_trend_p'] else 'N/A'
            }
        
        elif method == 'iv':
            formatted['工具变量法'] = {
                'IV系数': f"{result['iv_coef']:.4f}",
                '标准误': f"({result['iv_se']:.4f})",
                'P值': f"{result['iv_pvalue']:.4f}",
                '显著性': '***' if result['iv_pvalue'] < 0.01 else '**' if result['iv_pvalue'] < 0.05 else '*' if result['iv_pvalue'] < 0.1 else '',
                '第一阶段F统计量': f"{result['f_stat']:.2f}",
                '弱IV检验P值': f"{result['f_pvalue']:.4f}",
                '样本量': result['n_obs']
            }
        
        elif method == 'psm_att':
            formatted['PSM_ATT'] = {
                '简单ATT': f"{result['simple_att']:.4f}",
                '回归调整ATT': f"{result['reg_att']:.4f}",
                '标准误': f"({result['reg_se']:.4f})",
                'P值': f"{result['reg_pvalue']:.4f}",
                '显著性': '***' if result['reg_pvalue'] < 0.01 else '**' if result['reg_pvalue'] < 0.05 else '*' if result['reg_pvalue'] < 0.1 else '',
                '匹配后样本量': result['n_obs']
            }
        
        elif method == 'rdd':
            formatted['断点回归'] = {
                'LATE': f"{result['late']:.4f}",
                '标准误': f"({result['late_se']:.4f})",
                'P值': f"{result['late_pvalue']:.4f}",
                '显著性': '***' if result['late_pvalue'] < 0.01 else '**' if result['late_pvalue'] < 0.05 else '*' if result['late_pvalue'] < 0.1 else '',
                '样本量': result['n_obs'],
                '带宽': result['bandwidth'] or '全样本'
            }
        
        elif method == 'scm':
            effects_post = {k: v for k, v in result['effects'].items() if pd.notna(v)}
            avg_effect = np.mean(list(effects_post.values())) if effects_post else np.nan
            formatted['合成控制法'] = {
                '平均干预效应': f"{avg_effect:.4f}" if pd.notna(avg_effect) else 'N/A',
                '对照单元数': result['n_control_units'],
                '干预后时期数': len(effects_post)
            }
    
    return formatted</content>
<parameter name="filePath">D:\OpenCode-Projects\econometric-web\econometric_methods.py