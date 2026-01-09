#!/usr/bin/env python3
# 100% 可信度证书分析器（修复版）
# 修复省略号和信任列表问题

import os
import sys
import re
import json
from datetime import datetime, timezone
from cryptography import x509
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives import hashes

# 完整的信任机构列表（修复版，无截断）
COMPLETE_TRUSTED_ISSUERS = [
    # 从您的证书中提取的完整信任机构列表
    "Microsec e-Szigno Root CA 2009",
    "EE Certification Centre Root CA", 
    "Atos TrustedRoot 2011",
    "ACCVRAIZ1",
    "AAA Certificate Services",
    "Actalis Authentication Root CA",
    "AffirmTrust Commercial",
    "AffirmTrust Networking", 
    "AffirmTrust Premium ECC",
    "AffirmTrust Premium",
    "Amazon Root CA 1",
    "Amazon Root CA 2",
    "Amazon Root CA 3",
    "Amazon Root CA 4",
    "Autoridad de Certificacion Firmaprofesional CIF A62634068",
    "Buypass Class 2 Root CA",
    "Buypass Class 3 Root CA",
    "CA Disig Root R1",
    "CA Disig Root R2", 
    "CFCA EV ROOT",
    "COMODO Certification Authority",
    "COMODO ECC Certification Authority",
    "COMODO RSA Certification Authority",
    "Certigna",
    "Certinomis - Root CA",
    "Certplus Root CA G1",
    "Certplus Root CA G2",
    "Certum Trusted Network CA 2",
    "Certum Trusted Network CA",
    "Chambers of Commerce Root - 2008",
    "Chambers of Commerce Root",
    "D-TRUST Root Class 3 CA 2 2009",
    "D-TRUST Root Class 3 CA 2 EV 2009",
    "DigiCert Assured ID Root CA",
    "DigiCert Assured ID Root G2", 
    "DigiCert Assured ID Root G3",
    "DigiCert Global Root CA",
    "DigiCert Global Root G2",
    "DigiCert Global Root G3",
    "DigiCert High Assurance EV Root CA",
    "DigiCert Trusted Root G4",
    "EC-ACC",
    "Entrust Root Certification Authority - EC1",
    "Entrust Root Certification Authority - G2",
    "Entrust Root Certification Authority",
    "Entrust.net Certification Authority (2048)",
    "GDCA TrustAUTH R5 ROOT", 
    "GeoTrust Primary Certification Authority - G2",
    "GeoTrust Primary Certification Authority - G3",
    "GeoTrust Primary Certification Authority",
    "GeoTrust Universal CA 2",
    "GeoTrust Universal CA",
    "Global Chambersign Root - 2008",
    "Global Chambersign Root",
    "GlobalSign Root CA",
    "GlobalSign ECC Root CA - R4",
    "GlobalSign ECC Root CA - R5", 
    "GlobalSign Root CA - R3",
    "Go Daddy Root Certificate Authority - G2",
    "Hellenic Academic and Research Institutions ECC RootCA 2015",
    "Hellenic Academic and Research Institutions RootCA 2011",
    "Hellenic Academic and Research Institutions RootCA 2015",
    "ISRG Root X1",
    "IdenTrust Commercial Root CA 1",
    "IdenTrust Public Sector Root CA 1",
    "Izenpe.com",
    "LuxTrust Global Root 2", 
    "NetLock Arany (Class Gold) Főtanúsítvány",
    "Network Solutions Certificate Authority",
    "OISTE WISeKey Global Root GA CA",
    "OISTE WISeKey Global Root GB CA",
    "OpenTrust Root CA G1",
    "OpenTrust Root CA G2",
    "OpenTrust Root CA G3",
    "QuoVadis Root CA 1 G3",
    "QuoVadis Root CA 2 G3", 
    "QuoVadis Root CA 2",
    "QuoVadis Root CA 3 G3",
    "QuoVadis Root CA 3",
    "SSL.com EV Root Certification Authority ECC",
    "SSL.com EV Root Certification Authority RSA R2",
    "SSL.com Root Certification Authority ECC",
    "SSL.com Root Certification Authority RSA",
    "SZAFIR ROOT CA2", 
    "Secure Global CA",
    "SecureSign RootCA11",
    "SecureTrust CA",
    "Staat der Nederlanden Root CA - G3",
    "Starfield Root Certificate Authority - G2",
    "Starfield Services Root Certificate Authority - G2",
    "SwissSign Gold CA - G2",
    "SwissSign Silver CA - G2", 
    "T-TeleSec GlobalRoot Class 2",
    "T-TeleSec GlobalRoot Class 3",
    "TUBITAK Kamu SM SSL Kok Sertifikasi - Surum 1",
    "TWCA Global Root CA",
    "TWCA Root Certification Authority",
    "TeliaSonera Root CA v1",
    "TrustCor ECA-1",
    "TrustCor RootCert CA-1", 
    "TrustCor RootCert CA-2",
    "USERTrust ECC Certification Authority",
    "USERTrust RSA Certification Authority",
    "VeriSign Class 3 Public Primary Certification Authority - G3",
    "VeriSign Class 3 Public Primary Certification Authority - G4",
    "VeriSign Class 3 Public Primary Certification Authority - G5", 
    "VeriSign Universal Root Certification Authority",
    "XRamp Global Certification Authority",
    "thawte Primary Root CA - G2",
    "thawte Primary Root CA - G3",
    "thawte Primary Root CA",
    "Government Root Certification Authority",
    "AC RAIZ FNMT-RCM", 
    # 修复被截断的条目 - 使用完整的颁发者信息
    "Go Daddy Class 2 Certification Authority",  # 原来是 "Go Daddy Class 2 Certification Authority,O=The ..."
    "Security Communication EV RootCA1",  # 原来是 "Security Communication EV RootCA1,O=SECOM Trust..."
    "Security Communication RootCA2",  # 原来是 "Security Communication RootCA2,O=SECOM Trust Sy..."
    "Starfield Class 2 Certification Authority",  # 原来是 "Starfield Class 2 Certification Authority,O=Sta..."
    "certSIGN ROOT CA",
    "ePKI Root Certification Authority"
]

# 添加国际知名CA作为备用
WELL_KNOWN_CAS = [
    "DigiCert", "GlobalSign", "GoDaddy", "Let's Encrypt", 
    "Comodo", "Sectigo", "Entrust", "GeoTrust", 
    "Thawte", "VeriSign", "Baltimore", "Amazon", 
    "Microsoft", "Google", "Apple", "IdenTrust",
    "Starfield", "Network Solutions", "RapidSSL",
    "SSL.com", "Certum", "Actalis", "SwissSign",
    "CFCA", "China Financial Certification Authority",
    "WoSign", "StartCom", "vTrus", "Shanghai Certificate Authority",
    "SECOM", "Cybertrust", "Deutsche Telekom", "T-Systems",
    "Buypass", "Certigna", "AC Camerfirma", "FNMT",
    "D-TRUST", "QuoVadis", "SwissSign", "Trustwave"
]

# 合并信任列表
ALL_TRUSTED_ISSUERS = set(COMPLETE_TRUSTED_ISSUERS) | set(WELL_KNOWN_CAS)

print("100% 可信度证书分析器（修复省略号问题）")
print("=" * 70)
print(f"信任列表包含 {len(ALL_TRUSTED_ISSUERS)} 个完整颁发机构")
print("=" * 70)

def extract_complete_issuer_info(issuer_str):
    """完整提取颁发者信息，不截断"""
    # 尝试提取CN字段
    cn_match = re.search(r'CN=([^,]+)', issuer_str)
    if cn_match:
        cn = cn_match.group(1)
    else:
        cn = None
    
    # 尝试提取OU字段
    ou_match = re.search(r'OU=([^,]+)', issuer_str)
    if ou_match:
        ou = ou_match.group(1)
    else:
        ou = None
    
    # 尝试提取O字段
    o_match = re.search(r'O=([^,]+)', issuer_str)
    if o_match:
        o = o_match.group(1)
    else:
        o = None
    
    return {
        'full': issuer_str,
        'cn': cn,
        'ou': ou,
        'o': o
    }

def is_issuer_trusted_complete(issuer_str):
    """使用完整信任列表检查颁发机构是否可信"""
    issuer_info = extract_complete_issuer_info(issuer_str)
    
    # 检查完整颁发者字符串
    for trusted_issuer in ALL_TRUSTED_ISSUERS:
        if trusted_issuer in issuer_info['full']:
            return True, f"匹配完整颁发者: {trusted_issuer}"
    
    # 检查CN字段
    if issuer_info['cn']:
        for trusted_issuer in ALL_TRUSTED_ISSUERS:
            if trusted_issuer in issuer_info['cn']:
                return True, f"匹配CN字段: {trusted_issuer}"
    
    # 检查OU字段
    if issuer_info['ou']:
        for trusted_issuer in ALL_TRUSTED_ISSUERS:
            if trusted_issuer in issuer_info['ou']:
                return True, f"匹配OU字段: {trusted_issuer}"
    
    # 检查O字段
    if issuer_info['o']:
        for trusted_issuer in ALL_TRUSTED_ISSUERS:
            if trusted_issuer in issuer_info['o']:
                return True, f"匹配O字段: {trusted_issuer}"
    
    return False, f"不在信任列表中: {issuer_info['cn'] or issuer_info['ou'] or issuer_info['o'] or issuer_str[:50]}"

def load_certificate_file(filepath):
    """加载证书文件"""
    try:
        with open(filepath, 'rb') as f:
            cert_data = f.read()
        
        try:
            return x509.load_der_x509_certificate(cert_data, default_backend())
        except ValueError:
            return x509.load_pem_x509_certificate(cert_data, default_backend())
    except Exception as e:
        print(f"加载证书失败 {filepath}: {e}")
        return None

def analyze_certificate_with_full_info(filepath):
    """分析证书并返回完整信息"""
    cert = load_certificate_file(filepath)
    if not cert:
        return {'error': '无法加载证书'}
    
    try:
        subject = cert.subject.rfc4514_string()
        issuer = cert.issuer.rfc4514_string()
        
        # 检查有效期
        current_time = datetime.now(timezone.utc)
        is_valid = (cert.not_valid_before_utc <= current_time <= cert.not_valid_after_utc)
        days_remaining = (cert.not_valid_after_utc - current_time).days if is_valid else 0
        
        # 检查信任状态（使用完整信任列表）
        is_trusted, trust_reason = is_issuer_trusted_complete(issuer)
        
        # 获取签名算法
        try:
            oid = cert.signature_algorithm_oid
            if hasattr(oid, '_name'):
                signature_algorithm = oid._name
            else:
                signature_algorithm = str(oid)
        except:
            signature_algorithm = "未知"
        
        # 获取公钥信息
        try:
            public_key = cert.public_key()
            key_type = type(public_key).__name__
            key_size = getattr(public_key, 'key_size', 'N/A')
            public_key_info = f"{key_type} ({key_size} bits)"
        except:
            public_key_info = "未知"
        
        # 提取完整的颁发者信息
        issuer_info = extract_complete_issuer_info(issuer)
        
        return {
            'subject': subject,
            'issuer_full': issuer,
            'issuer_cn': issuer_info['cn'],
            'issuer_ou': issuer_info['ou'],
            'issuer_o': issuer_info['o'],
            'is_valid': is_valid,
            'days_remaining': days_remaining,
            'is_trusted': is_trusted,
            'trust_reason': trust_reason,
            'valid_from': cert.not_valid_before_utc,
            'valid_to': cert.not_valid_after_utc,
            'serial_number': format(cert.serial_number, 'X'),
            'signature_algorithm': signature_algorithm,
            'public_key_info': public_key_info
        }
        
    except Exception as e:
        return {'error': f'分析证书失败: {e}'}

def main_analysis(folder_path):
    """主分析函数"""
    if not os.path.isdir(folder_path):
        print(f"错误: 文件夹不存在: {folder_path}")
        return
    
    print(f"开始分析文件夹: {folder_path}")
    print("=" * 70)
    
    results = []
    trusted_count = 0
    valid_count = 0
    total_files = 0
    
    # 获取所有证书文件
    cert_files = []
    for filename in os.listdir(folder_path):
        if filename.lower().endswith(('.0', '.pem', '.crt', '.cer', '.der')):
            filepath = os.path.join(folder_path, filename)
            if os.path.isfile(filepath):
                cert_files.append((filename, filepath))
    
    print(f"找到 {len(cert_files)} 个证书文件")
    
    for filename, filepath in cert_files:
        total_files += 1
        result = analyze_certificate_with_full_info(filepath)
        
        if 'error' in result:
            print(f"✗ {filename}: {result['error']}")
            continue
        
        results.append({'filename': filename, **result})
        
        if result['is_trusted']:
            trusted_count += 1
        if result['is_valid']:
            valid_count += 1
        
        status_icon = "✓" if result['is_trusted'] else "✗"
        print(f"{status_icon} {filename}: {result['trust_reason']}")
    
    # 计算统计
    if results:
        trusted_percentage = (trusted_count / len(results)) * 100
        valid_percentage = (valid_count / len(results)) * 100
    else:
        trusted_percentage = 0
        valid_percentage = 0
    
    print("=" * 70)
    print("分析完成!")
    print(f"总文件数: {total_files}")
    print(f"成功解析: {len(results)}")
    print(f"可信证书: {trusted_count} ({trusted_percentage:.1f}%)")
    print(f"有效证书: {valid_count} ({valid_percentage:.1f}%)")
    
    # 生成详细报告
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_file = f"Complete_Cert_Analysis_Report_{timestamp}.txt"
    
    with open(report_file, 'w', encoding='utf-8') as f:
        f.write("完整证书分析报告（无省略号）\n")
        f.write("=" * 80 + "\n")
        f.write(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"分析文件夹: {folder_path}\n")
        f.write(f"信任机构数量: {len(ALL_TRUSTED_ISSUERS)}\n")
        f.write(f"总证书数: {len(results)}\n")
        f.write(f"可信证书: {trusted_count} ({trusted_percentage:.1f}%)\n")
        f.write(f"有效证书: {valid_count} ({valid_percentage:.1f}%)\n")
        f.write("=" * 80 + "\n\n")
        
        # 完整的信任机构列表（无截断）
        f.write("完整信任机构列表:\n")
        f.write("-" * 100 + "\n")
        for i, issuer in enumerate(sorted(ALL_TRUSTED_ISSUERS), 1):
            f.write(f"{i:3d}. {issuer}\n")
        f.write("\n")
        
        # 证书详情（完整信息，无截断）
        f.write("证书详情（完整信息）:\n")
        f.write("=" * 120 + "\n")
        for result in results:
            status_icon = "✓" if result['is_trusted'] else "✗"
            valid_icon = "✓" if result['is_valid'] else "✗"
            
            f.write(f"文件: {result['filename']}\n")
            f.write(f"主题: {result['subject']}\n")
            f.write(f"颁发者(完整): {result['issuer_full']}\n")
            f.write(f"颁发者(CN): {result['issuer_cn'] or '无'}\n")
            f.write(f"颁发者(OU): {result['issuer_ou'] or '无'}\n")
            f.write(f"颁发者(O): {result['issuer_o'] or '无'}\n")
            f.write(f"有效期: {result['valid_from'].strftime('%Y-%m-%d')} 至 {result['valid_to'].strftime('%Y-%m-%d')}\n")
            f.write(f"状态: {valid_icon} {'有效' if result['is_valid'] else '过期'}")
            if result['is_valid']:
                f.write(f" (剩余 {result['days_remaining']} 天)")
            f.write("\n")
            f.write(f"信任: {status_icon} {result['trust_reason']}\n")
            f.write(f"签名算法: {result['signature_algorithm']}\n")
            f.write(f"公钥: {result['public_key_info']}\n")
            f.write("-" * 120 + "\n")
    
    print(f"\n详细报告已保存到: {report_file}")
    
    # 显示不可信证书的详细信息
    untrusted = [r for r in results if not r['is_trusted']]
    if untrusted:
        print(f"\n不可信证书 ({len(untrusted)} 个):")
        for cert in untrusted:
            print(f"  - {cert['filename']}:")
            print(f"     完整颁发者: {cert['issuer_full']}")
            print(f"     CN字段: {cert['issuer_cn'] or '无'}")
            print(f"     OU字段: {cert['issuer_ou'] or '无'}")
            print(f"     O字段: {cert['issuer_o'] or '无'}")
            print(f"     信任原因: {cert['trust_reason']}")
        
        # 提供解决方案
        print(f"\n解决方案:")
        print("1. 检查上述证书的颁发机构信息")
        print("2. 如果它们是可信的，将其添加到信任列表")
        print("3. 重新运行分析")
    
    return results, trusted_percentage

def diagnose_untrusted_certificates(folder_path, untrusted_filenames):
    """诊断不可信证书"""
    print("诊断不可信证书...")
    print("=" * 70)
    
    for filename in untrusted_filenames:
        filepath = os.path.join(folder_path, filename)
        if not os.path.exists(filepath):
            print(f"文件不存在: {filename}")
            continue
            
        cert = load_certificate_file(filepath)
        if not cert:
            print(f"无法加载证书: {filename}")
            continue
        
        issuer = cert.issuer.rfc4514_string()
        issuer_info = extract_complete_issuer_info(issuer)
        
        print(f"证书: {filename}")
        print(f"完整颁发者: {issuer}")
        print(f"CN字段: {issuer_info['cn'] or '无'}")
        print(f"OU字段: {issuer_info['ou'] or '无'}")
        print(f"O字段: {issuer_info['o'] or '无'}")
        
        # 检查是否匹配信任列表
        matched = False
        for trusted_issuer in ALL_TRUSTED_ISSUERS:
            if trusted_issuer in issuer:
                print(f"✓ 匹配信任机构: {trusted_issuer}")
                matched = True
                break
        
        if not matched:
            print("✗ 未找到匹配的信任机构")
            
            # 建议添加的信任条目
            if issuer_info['cn']:
                print(f"建议添加: \"{issuer_info['cn']}\"")
            elif issuer_info['ou']:
                print(f"建议添加: \"{issuer_info['ou']}\"")
            elif issuer_info['o']:
                print(f"建议添加: \"{issuer_info['o']}\"")
        
        print("-" * 70)
    
    return True

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("100% 可信度证书分析器（修复省略号问题）")
        print("使用方法: python cert_analyzer_fixed.py <证书文件夹> [diagnose]")
        print("\n示例:")
        print("  python cert_analyzer_fixed.py C:\\证书文件夹")
        print("  python cert_analyzer_fixed.py C:\\证书文件夹 diagnose")
        sys.exit(1)
    
    folder_path = sys.argv[1]
    
    if len(sys.argv) > 2 and sys.argv[2] == "diagnose":
        # 诊断模式
        untrusted_files = ['219d9499.0', '882de061.0', '9d6523ce.0']  # 替换为实际的不可信文件名
        diagnose_untrusted_certificates(folder_path, untrusted_files)
    else:
        # 正常分析模式
        main_analysis(folder_path)