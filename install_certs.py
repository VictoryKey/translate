import os
import sys
import tempfile
import subprocess
import shutil
from cryptography import x509
from cryptography.hazmat.backends import default_backend
import winreg
import ctypes
import argparse

def is_admin():
    """检查是否以管理员权限运行"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    """以管理员权限重新运行脚本"""
    if not is_admin():
        script = os.path.abspath(sys.argv[0])
        params = ' '.join([f'"{arg}"' for arg in sys.argv[1:]])
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{script}" {params}', None, 1)
        sys.exit(0)

def get_certificate_files(folder_path):
    """获取文件夹中的所有证书文件"""
    cert_extensions = ('.crt', '.cer', '.pem', '.der', '.0')
    certificate_files = []
    
    for filename in os.listdir(folder_path):
        if filename.lower().endswith(cert_extensions):
            filepath = os.path.join(folder_path, filename)
            if os.path.isfile(filepath):
                certificate_files.append(filepath)
    
    return certificate_files

def load_certificate(filepath):
    """加载证书文件"""
    try:
        with open(filepath, 'rb') as f:
            cert_data = f.read()
        
        # 尝试PEM格式
        try:
            cert = x509.load_pem_x509_certificate(cert_data, default_backend())
            return cert, 'PEM'
        except ValueError:
            # 尝试DER格式
            try:
                cert = x509.load_der_x509_certificate(cert_data, default_backend())
                return cert, 'DER'
            except ValueError:
                print(f"  错误: 无法解析证书格式: {filepath}")
                return None, None
                
    except Exception as e:
        print(f"  错误: 读取证书文件失败: {filepath} - {e}")
        return None, None

def get_certificate_info(cert):
    """获取证书基本信息"""
    try:
        subject = cert.subject.rfc4514_string()
        issuer = cert.issuer.rfc4514_string()
        serial = format(cert.serial_number, 'X')
        not_before = cert.not_valid_before
        not_after = cert.not_valid_after
        
        return {
            'subject': subject,
            'issuer': issuer,
            'serial': serial,
            'not_before': not_before,
            'not_after': not_after,
            'is_self_signed': cert.issuer == cert.subject
        }
    except Exception as e:
        print(f"  错误: 获取证书信息失败: {e}")
        return None

def is_certificate_installed(cert_info):
    """检查证书是否已安装"""
    try:
        # 使用certutil检查证书是否已存在
        result = subprocess.run(
            ['certutil', '-verifystore', 'Root'],
            capture_output=True, text=True, encoding='utf-8', errors='ignore'
        )
        
        # 检查主题是否已存在
        if cert_info['subject'] in result.stdout:
            return True
        
        # 检查序列号是否已存在
        if f"Serial Number: {cert_info['serial']}" in result.stdout:
            return True
            
        return False
    except Exception as e:
        print(f"  警告: 检查证书安装状态失败: {e}")
        return False

def install_certificate(filepath, cert_info):
    """安装单个证书到受信任根证书存储"""
    try:
        # 方法1: 使用certutil（推荐）
        result = subprocess.run(
            ['certutil', '-addstore', 'Root', filepath],
            capture_output=True, text=True, encoding='utf-8', errors='ignore'
        )
        
        if "ERROR" in result.stdout or "失败" in result.stdout:
            # 方法2: 使用PowerShell（备用方法）
            ps_command = f'''
                $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2("{filepath}")
                $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "LocalMachine")
                $store.Open("ReadWrite")
                $store.Add($cert)
                $store.Close()
            '''
            
            result_ps = subprocess.run(
                ['powershell', '-Command', ps_command],
                capture_output=True, text=True, encoding='utf-8', errors='ignore'
            )
            
            if result_ps.returncode != 0:
                print(f"  PowerShell安装失败: {result_ps.stderr}")
                return False
            else:
                print("  ✓ 使用PowerShell安装成功")
                return True
        else:
            print("  ✓ 使用certutil安装成功")
            return True
            
    except Exception as e:
        print(f"  错误: 安装证书失败: {e}")
        return False

def export_certificate_to_der(cert, output_path):
    """将证书导出为DER格式（certutil偏好格式）"""
    try:
        with open(output_path, 'wb') as f:
            f.write(cert.public_bytes(encoding=serialization.Encoding.DER))
        return True
    except Exception as e:
        print(f"  错误: 导出证书为DER格式失败: {e}")
        return False

def install_certificates_batch(folder_path, force_reinstall=False):
    """批量安装证书"""
    print("=" * 70)
    print("Windows 证书批量导入工具")
    print("=" * 70)
    
    # 检查文件夹是否存在
    if not os.path.isdir(folder_path):
        print(f"错误: 文件夹不存在: {folder_path}")
        return False
    
    # 获取所有证书文件
    cert_files = get_certificate_files(folder_path)
    if not cert_files:
        print(f"错误: 在文件夹中未找到证书文件: {folder_path}")
        print("支持的格式: .crt, .cer, .pem, .der, .0")
        return False
    
    print(f"找到 {len(cert_files)} 个证书文件")
    print("开始分析证书...")
    print("-" * 70)
    
    # 创建临时目录
    temp_dir = tempfile.mkdtemp()
    
    installed_count = 0
    skipped_count = 0
    failed_count = 0
    
    try:
        for i, cert_file in enumerate(cert_files, 1):
            filename = os.path.basename(cert_file)
            print(f"{i}/{len(cert_files)} 处理: {filename}")
            
            # 加载证书
            cert, cert_format = load_certificate(cert_file)
            if not cert:
                failed_count += 1
                continue
            
            # 获取证书信息
            cert_info = get_certificate_info(cert)
            if not cert_info:
                failed_count += 1
                continue
            
            # 显示证书信息
            print(f"  主题: {cert_info['subject'][:80]}...")
            print(f"  颁发者: {cert_info['issuer'][:80]}...")
            print(f"  序列号: {cert_info['serial']}")
            print(f"  有效期: {cert_info['not_before'].strftime('%Y-%m-%d')} 至 {cert_info['not_after'].strftime('%Y-%m-%d')}")
            print(f"  自签名: {'是' if cert_info['is_self_signed'] else '否'}")
            
            # 检查是否已安装
            if not force_reinstall and is_certificate_installed(cert_info):
                print("  ⚠ 证书已存在，跳过安装")
                skipped_count += 1
                continue
            
            # 对于PEM格式，转换为DER格式（certutil处理PEM有时有问题）
            if cert_format == 'PEM':
                der_file = os.path.join(temp_dir, f"temp_{i}.der")
                try:
                    from cryptography.hazmat.primitives import serialization
                    with open(der_file, 'wb') as f:
                        f.write(cert.public_bytes(encoding=serialization.Encoding.DER))
                    install_file = der_file
                except Exception as e:
                    print(f"  ⚠ PEM转DER失败，使用原文件: {e}")
                    install_file = cert_file
            else:
                install_file = cert_file
            
            # 安装证书
            if install_certificate(install_file, cert_info):
                installed_count += 1
                print("  ✓ 安装成功")
            else:
                failed_count += 1
                print("  ✗ 安装失败")
            
            print("-" * 70)
        
        # 显示安装结果
        print("\n安装完成!")
        print("=" * 70)
        print(f"成功安装: {installed_count} 个证书")
        print(f"跳过已存在: {skipped_count} 个证书")
        print(f"安装失败: {failed_count} 个证书")
        print("=" * 70)
        
        if installed_count > 0:
            print("\n注意: 某些应用程序可能需要重启才能识别新安装的证书")
            print("建议重启计算机以确保所有应用程序都能识别新证书")
        
        return failed_count == 0
        
    finally:
        # 清理临时文件
        try:
            shutil.rmtree(temp_dir)
        except:
            pass

def list_installed_certificates():
    """列出已安装的根证书"""
    try:
        result = subprocess.run(
            ['certutil', '-store', 'Root'],
            capture_output=True, text=True, encoding='utf-8', errors='ignore'
        )
        
        if result.returncode == 0:
            print("已安装的根证书:")
            print("=" * 70)
            print(result.stdout)
        else:
            print("获取证书列表失败")
            
    except Exception as e:
        print(f"错误: 无法列出证书: {e}")

def remove_certificate_by_serial(serial_number):
    """根据序列号删除证书"""
    try:
        result = subprocess.run(
            ['certutil', '-delstore', 'Root', serial_number],
            capture_output=True, text=True, encoding='utf-8', errors='ignore'
        )
        
        if result.returncode == 0:
            print(f"✓ 证书序列号 {serial_number} 删除成功")
            return True
        else:
            print(f"✗ 删除证书失败: {result.stderr}")
            return False
            
    except Exception as e:
        print(f"错误: 删除证书失败: {e}")
        return False

def main():
    """主函数"""
    parser = argparse.ArgumentParser(description='Windows证书批量导入工具')
    parser.add_argument('folder', nargs='?', help='包含证书文件的文件夹路径')
    parser.add_argument('-f', '--force', action='store_true', help='强制重新安装已存在的证书')
    parser.add_argument('-l', '--list', action='store_true', help='列出已安装的根证书')
    parser.add_argument('-r', '--remove', help='根据序列号删除证书')
    
    args = parser.parse_args()
    
    # 检查管理员权限
    if not is_admin():
        print("此操作需要管理员权限")
        run_as_admin()
        return
    
    if args.list:
        list_installed_certificates()
        return
    
    if args.remove:
        remove_certificate_by_serial(args.remove)
        return
    
    if not args.folder:
        print("请指定证书文件夹路径")
        print("使用方法:")
        print("  python install_certs.py <证书文件夹路径>")
        print("  python install_certs.py -l  # 列出已安装证书")
        print("  python install_certs.py -r <序列号>  # 删除证书")
        print("  python install_certs.py -f <证书文件夹路径>  # 强制重新安装")
        return
    
    folder_path = args.folder
    if not os.path.isdir(folder_path):
        print(f"错误: 文件夹不存在: {folder_path}")
        return
    
    print("警告: 此操作将向系统受信任的根证书存储添加证书")
    print("请确保这些证书来自可信来源!")
    print("")
    confirmation = input("是否继续? (y/N): ")
    
    if confirmation.lower() not in ['y', 'yes']:
        print("操作已取消")
        return
    
    # 执行批量安装
    success = install_certificates_batch(folder_path, args.force)
    
    if success:
        print("证书安装完成!")
    else:
        print("证书安装过程中遇到错误，请检查上方日志")

if __name__ == "__main__":
    main()