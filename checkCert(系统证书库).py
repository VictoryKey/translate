import os
import sys
import warnings
from datetime import datetime, timezone
from cryptography import x509
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives import hashes
import cryptography.exceptions
import subprocess
import certifi
import ssl

# 抑制特定的弃用警告
warnings.filterwarnings('ignore', category=cryptography.utils.CryptographyDeprecationWarning)

def load_system_trust_store():
    """加载系统信任的根证书库"""
    trust_store = {}
    loaded_certs = 0
    
    print("尝试加载系统根证书库...")
    
    # 首先尝试 certifi
    try:
        certifi_path = certifi.where()
        if os.path.exists(certifi_path):
            print(f"✓ 使用 certifi 证书: {certifi_path}")
            certs = load_certificates_from_file(certifi_path)
            for cert in certs:
                fingerprint = cert.fingerprint(hashes.SHA256()).hex()
                trust_store[fingerprint] = cert
                loaded_certs += 1
            print(f"✓ 从 certifi 加载 {len(certs)} 个证书")
    except Exception as e:
        print(f"⚠ certifi 加载失败: {e}")
    
    # Windows 系统路径
    if os.name == 'nt':
        windows_paths = [
            os.path.join(os.environ.get('SystemRoot', 'C:\\Windows'), 'System32', 'CertSrv', 'CertEnroll'),
            os.path.join(os.environ.get('SystemRoot', 'C:\\Windows'), 'System32', 'certs'),
        ]
        
        for path in windows_paths:
            if os.path.exists(path):
                try:
                    if os.path.isfile(path):
                        certs = load_certificates_from_file(path)
                    else:
                        certs = load_certificates_from_directory(path)
                    
                    for cert in certs:
                        fingerprint = cert.fingerprint(hashes.SHA256()).hex()
                        trust_store[fingerprint] = cert
                        loaded_certs += 1
                    
                    print(f"✓ 从系统路径加载: {path} ({len(certs)} 个证书)")
                except Exception as e:
                    print(f"  警告: 加载 {path} 失败: {e}")
    
    if loaded_certs == 0:
        print("⚠ 警告: 无法加载任何系统根证书")
    else:
        print(f"✓ 成功加载 {loaded_certs} 个系统信任的根证书")
    
    return trust_store

def load_certificates_from_file(filepath):
    """从文件加载证书"""
    certificates = []
    try:
        with open(filepath, 'rb') as f:
            data = f.read()
        
        # 处理PEM格式（可能包含多个证书）
        if b'-----BEGIN CERTIFICATE-----' in data:
            # 分割PEM证书
            pem_blocks = data.split(b'-----BEGIN CERTIFICATE-----')
            for block in pem_blocks[1:]:  # 跳过第一个空块
                if b'-----END CERTIFICATE-----' not in block:
                    continue
                
                cert_pem = b'-----BEGIN CERTIFICATE-----' + block.split(b'-----END CERTIFICATE-----')[0] + b'-----END CERTIFICATE-----'
                
                try:
                    cert = x509.load_pem_x509_certificate(cert_pem, default_backend())
                    certificates.append(cert)
                except Exception as e:
                    continue
        else:
            # 尝试作为单个证书处理
            try:
                cert = x509.load_pem_x509_certificate(data, default_backend())
                certificates.append(cert)
            except ValueError:
                try:
                    cert = x509.load_der_x509_certificate(data, default_backend())
                    certificates.append(cert)
                except ValueError:
                    pass
                
    except Exception as e:
        print(f"  读取文件失败 {filepath}: {e}")
    
    return certificates

def load_certificates_from_directory(dirpath):
    """从目录加载证书"""
    certificates = []
    if not os.path.isdir(dirpath):
        return certificates
    
    for filename in os.listdir(dirpath):
        if not filename.lower().endswith(('.pem', '.crt', '.cer', '.der', '.0')):
            continue
        
        filepath = os.path.join(dirpath, filename)
        if os.path.isfile(filepath):
            try:
                certs = load_certificates_from_file(filepath)
                certificates.extend(certs)
            except Exception as e:
                print(f"  加载 {filename} 失败: {e}")
    
    return certificates

def safe_serial_number(serial_number):
    """安全处理序列号"""
    if serial_number < 0:
        # 将负数转换为正数
        max_bits = 159
        max_serial = (1 << max_bits) - 1
        return serial_number & max_serial
    return serial_number

def check_certificate_file(filepath):
    """检查单个证书文件"""
    try:
        with open(filepath, 'rb') as f:
            cert_data = f.read()
        
        # 尝试不同格式
        try:
            cert = x509.load_der_x509_certificate(cert_data, default_backend())
        except ValueError:
            try:
                cert = x509.load_pem_x509_certificate(cert_data, default_backend())
            except ValueError as e:
                print(f"无法解析证书 {os.path.basename(filepath)}: {e}")
                return None
        
        return cert
        
    except Exception as e:
        print(f"读取证书文件失败 {filepath}: {e}")
        return None

def check_trust_status(cert, trust_store):
    """检查证书信任状态"""
    try:
        issuer = cert.issuer.rfc4514_string()
        subject = cert.subject.rfc4514_string()
        
        # 检查是否为自签名证书
        if cert.issuer == cert.subject:
            cert_fingerprint = cert.fingerprint(hashes.SHA256()).hex()
            if cert_fingerprint in trust_store:
                return True, "系统信任的自签名根证书"
            else:
                return False, "自签名证书，但不在系统信任库中"
        else:
            # 检查颁发者是否在信任库中
            issuer_found = False
            for trusted_cert in trust_store.values():
                if trusted_cert.subject == cert.issuer:
                    issuer_found = True
                    break
            
            if issuer_found:
                return True, f"颁发者在系统信任库中"
            else:
                return False, f"颁发者不在系统信任库中"
    except Exception as e:
        return False, f"信任检查错误: {e}"

def get_signature_algorithm(cert):
    """获取签名算法信息"""
    try:
        oid = cert.signature_algorithm_oid
        if hasattr(oid, '_name'):
            return oid._name
        return str(oid)
    except:
        return "未知"

def get_public_key_info(cert):
    """获取公钥信息"""
    try:
        public_key = cert.public_key()
        key_type = type(public_key).__name__
        key_size = getattr(public_key, 'key_size', 'N/A')
        return f"{key_type} ({key_size} bits)"
    except:
        return "未知"

def analyze_certificates(folder_path, output_file):
    """分析文件夹中的所有证书文件 - 这是缺失的函数"""
    # 加载系统信任库
    trust_store = load_system_trust_store()
    
    results = []
    processed_files = 0
    successful_parses = 0
    
    print(f"\n开始分析目录: {folder_path}")
    
    # 遍历文件夹
    for filename in os.listdir(folder_path):
        # 支持常见的证书文件扩展名
        if not filename.lower().endswith(('.0', '.pem', '.crt', '.cer', '.der')):
            continue
            
        filepath = os.path.join(folder_path, filename)
        if not os.path.isfile(filepath):
            continue
            
        processed_files += 1
        cert = check_certificate_file(filepath)
        
        if cert is None:
            continue
            
        successful_parses += 1
        
        try:
            # 使用UTC时间属性
            not_before = cert.not_valid_before_utc
            not_after = cert.not_valid_after_utc
            current_time = datetime.now(timezone.utc)
            
            # 检查有效期
            is_valid = not_before <= current_time <= not_after
            days_remaining = (not_after - current_time).days if is_valid else 0
            
            # 安全处理序列号
            serial_number = safe_serial_number(cert.serial_number)
            
            # 检查信任状态
            is_trusted, trust_reason = check_trust_status(cert, trust_store)
            
            results.append({
                'filename': filename,
                'subject': cert.subject.rfc4514_string(),
                'issuer': cert.issuer.rfc4514_string(),
                'serial_number': f"{serial_number:x}",
                'valid_from': not_before,
                'valid_to': not_after,
                'is_valid': is_valid,
                'days_remaining': days_remaining,
                'is_trusted': is_trusted,
                'trust_reason': trust_reason,
                'signature_algorithm': get_signature_algorithm(cert),
                'public_key_info': get_public_key_info(cert)
            })
            
        except Exception as e:
            print(f"分析证书 {filename} 时出错: {e}")
            successful_parses -= 1
    
    # 生成报告
    generate_report(results, output_file, trust_store, processed_files, successful_parses)
    
    return results

def generate_report(results, output_file, trust_store, processed_files, successful_parses):
    """生成分析报告"""
    with open(output_file, 'w', encoding='utf-8') as f:
        # 头部信息
        f.write("=" * 70 + "\n")
        f.write("证书分析报告\n")
        f.write("=" * 70 + "\n")
        f.write(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"系统信任证书: {len(trust_store)} 个\n")
        f.write(f"处理文件数: {processed_files} 个\n")
        f.write(f"成功解析: {successful_parses} 个\n")
        f.write("=" * 70 + "\n\n")
        
        # 统计信息
        if results:
            valid_certs = sum(1 for r in results if r['is_valid'])
            trusted_certs = sum(1 for r in results if r['is_trusted'])
            
            f.write("统计摘要:\n")
            f.write(f"- 有效证书: {valid_certs}/{len(results)} ({valid_certs/len(results)*100:.1f}%)\n")
            f.write(f"- 可信证书: {trusted_certs}/{len(results)} ({trusted_certs/len(results)*100:.1f}%)\n\n")
        else:
            f.write("统计摘要: 无有效证书数据\n\n")
        
        # 详细结果
        f.write("证书详细信息:\n")
        f.write("=" * 100 + "\n")
        
        if results:
            for i, result in enumerate(results, 1):
                status_icon = "✓" if result['is_valid'] else "✗"
                trusted_icon = "✓" if result['is_trusted'] else "✗"
                status_text = "有效" if result['is_valid'] else "过期"
                
                f.write(f"{i}. {result['filename']}\n")
                f.write(f"   主题: {result['subject']}\n")
                f.write(f"   颁发者: {result['issuer']}\n")
                f.write(f"   序列号: {result['serial_number']}\n")
                f.write(f"   有效期: {result['valid_from'].strftime('%Y-%m-%d')} 至 {result['valid_to'].strftime('%Y-%m-%d')}\n")
                f.write(f"   状态: {status_icon} {status_text}")
                if result['is_valid']:
                    f.write(f" (剩余 {result['days_remaining']} 天)")
                f.write("\n")
                f.write(f"   信任: {trusted_icon} {result['trust_reason']}\n")
                f.write(f"   签名算法: {result['signature_algorithm']}\n")
                f.write(f"   公钥: {result['public_key_info']}\n")
                f.write("-" * 100 + "\n")
        else:
            f.write("无证书数据可显示\n")
    
    # 控制台输出摘要
    print(f"\n✓ 分析完成!")
    print(f"✓ 报告文件: {output_file}")
    print(f"✓ 处理统计: {successful_parses}/{processed_files} 个文件成功解析")
    
    if results:
        valid_pct = sum(1 for r in results if r['is_valid']) / len(results) * 100
        trusted_pct = sum(1 for r in results if r['is_trusted']) / len(results) * 100
        print(f"✓ 有效证书: {valid_pct:.1f}% | 可信证书: {trusted_pct:.1f}%")

def main():
    if len(sys.argv) != 2:
        print("证书分析工具")
        print("使用方法: python checkCert.py <证书文件夹路径>")
        print("\n示例:")
        print("  python checkCert.py C:\\Users\\Bamb00\\Desktop\\crt\\cacerts")
        print("\n支持的文件格式: .0, .pem, .crt, .cer, .der")
        sys.exit(1)
    
    folder_path = sys.argv[1]
    
    if not os.path.isdir(folder_path):
        print(f"错误: 不是有效的文件夹: {folder_path}")
        sys.exit(1)
    
    # 显示系统信息
    print("=" * 60)
    print("证书分析工具 - 系统信息")
    print("=" * 60)
    print(f"操作系统: {os.name}")
    print(f"Python版本: {sys.version.split()[0]}")
    print(f"证书文件夹: {folder_path}")
    
    # 显示certifi信息
    try:
        certifi_path = certifi.where()
        print(f"certifi证书路径: {certifi_path}")
    except Exception as e:
        print(f"certifi信息: 不可用 ({e})")
    
    print("=" * 60)
    
    # 设置输出文件名
    folder_name = os.path.basename(folder_path.rstrip('/\\'))
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"CertResult_{folder_name}_{timestamp}.txt"
    
    # 执行分析
    try:
        start_time = datetime.now()
        analyze_certificates(folder_path, output_file)
        end_time = datetime.now()
        print(f"\n分析耗时: {(end_time - start_time).total_seconds():.2f}秒")
        
    except KeyboardInterrupt:
        print("\n用户中断操作")
    except Exception as e:
        print(f"分析过程中发生错误: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()