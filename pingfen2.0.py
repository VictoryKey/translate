# 输入技术细节掌握（TD）分数
print("请输入技术细节掌握（TD）分数，TD得分 分别是 \n低 1分\n中 3分\n高 10分 ")
TD = float(input())

# 输入影响车辆工况（VC）分数
print("请输入影响车辆工况（VC）分数，VC得分 分别是 \n未定义 0分\n静止 1分\n低速 2分\n中速 4分\n较高 7分\n高速 10分  ")
VC = float(input())

# 输入漏洞通用性（VV）分数
print("请输入漏洞通用性（VV），VV得分 分别是 \n单车攻击 1分\n单一车型攻击 7分\n多车型攻击 10分 ")
VV = float(input())

# 输入利用窗口（WD）分数
print("请输入利用窗口（WD）分数，WD得分 分别是 \n物理接触 5分\n本地 6分\n近源 8分\n远程 10分 ")
WD = float(input())

# 输入所需知识技能（KS）分数
print("请输入所需知识技能（KS）分数，KS得分 分别是 \n多领域专家 1分\n汽车安全专家 2分\n熟练操作者 5分\n业余者 10分 ")
KS = float(input())

# 输入所需设备（EM）分数
print("请输入所需设备（EM）分数，EM得分 分别是 \n多种定制或专有的硬件和软件 1分\n定制或专有的硬件设备和软件 3分\n公开的专用硬件设备和软件 7分\n公开的硬件设备和软件 10分 ")
EM = float(input())

# 输入影响范围（IS）分数
print("请输入影响范围（IS）分数，IS得分 分别是 \n单一 1分\n多个 7分\n几乎全部 10分 ")
AS = float(input())

# 输入人身安全（PS）分数
print("请输入人身安全（PS）分数，PS得分 分别是\n无 0分\n轻度伤害 3分\n严重受伤 7分\n生命威胁 10分 ")
PS = float(input())

# 输入财产安全（PP）分数
print("请输入财产安全（PP）分数，PP得分 分别是\n无 0分\n低 2分\n中 7分\n高 10分 ")
PP = float(input())

# 输入操作影响（OA）分数
print("请输入操作影响（OA）分数，OA得分 分别是\n无 0分\n低 2分\n中 4分\n高 10分")
OA = float(input())

# 输入隐私安全（PA）分数
print("请输入隐私安全（PA）分数，PA得分 分别是\n无 0分\n低 1分\n中 6分\n高 10分")
PA = float(input())

# 输入公共安全及法规（PR）分数
print("请输入公共安全及法规（PR）PR得分 分别是\n无 0分\n低 2分\n中 6分\n高 10分 ")
PR = float(input())

# 计算场景参数（SP）、威胁参数（TP）和影响参数（IP）
SP = 1.951 * TD + 1.341 * VC + 1.341 * VV
TP = 1.951 * WD + 0.122 * KS + 0.732 * EM + 2.561 * AS
IP = 3.789 * PS + 1.684 * PP + 1.158 * OA + 1.158 * PA + 2.211 * PR

# 计算攻击等级评分和影响等级评分
def calculate_level(score):
    if score <= 20:
        return 1
    elif score <= 50:
        return 2
    elif score <= 75:
        return 3
    else:
        return 4

attack_level = calculate_level(SP + TP)
impact_level = calculate_level(IP)

# 根据评分矩阵确定漏洞等级
vulnerability_matrix = [
    ["低危", "低危", "中危", "中危"],
    ["低危", "中危", "中危", "高危"],
    ["中危", "中危", "高危", "高危"],
    ["中危", "高危", "高危", "高危"]
]

vulnerability_level = vulnerability_matrix[attack_level - 1][impact_level - 1]

# 输出最终得分和漏洞等级
print("最终得分 ：\n SP = " + str(SP) + " \n TP = " + str(TP) + " \n IP = " + str(IP))
print("攻击等级评分: " + str(attack_level))
print("影响等级评分: " + str(impact_level))
print("汽车漏洞等级: " + vulnerability_level)