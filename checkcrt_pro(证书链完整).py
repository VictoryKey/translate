import os
import sys
import warnings
from datetime import datetime, timezone
from cryptography import x509
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives import hashes
import collections

# 抑制警告
warnings.filterwarnings('ignore')

class CertificateAnalyzer:
    def __init__(self, certs_folder):
        self.certs_folder = certs_folder
        self.trust_store = {}
        self.all_certs = []
        self.cert_chains = {}
        
    def load_all_certificates(self):
        """加载所有证书"""
        print("加载证书文件...")
        
        for filename in os.listdir(self.certs_folder):
            if not filename.lower().endswith(('.0', '.pem', '.crt', '.cer', '.der')):
                continue
                
            filepath = os.path.join(self.certs_folder, filename)
            cert = self.load_certificate_file(filepath)
            if cert:
                self.all_certs.append((filename, cert))
        
        print(f"✓ 加载了 {len(self.all_certs)} 个证书")
        return self.all_certs
    
    def load_certificate_file(self, filepath):
        """加载单个证书文件"""
        try:
            with open(filepath, 'rb') as f:
                cert_data = f.read()
            
            try:
                return x509.load_der_x509_certificate(cert_data, default_backend())
            except ValueError:
                return x509.load_pem_x509_certificate(cert_data, default_backend())
        except:
            return None
    
    def build_certificate_chains(self):
        """构建所有可能的证书链"""
        print("构建证书链...")
        
        # 首先识别根证书（自签名）
        root_certs = []
        intermediate_certs = []
        leaf_certs = []
        
        for filename, cert in self.all_certs:
            if cert.issuer == cert.subject:
                root_certs.append((filename, cert))
            else:
                # 检查是否为中间证书（被其他证书引用为颁发者）
                is_intermediate = any(
                    c.issuer == cert.subject for _, c in self.all_certs 
                    if c.issuer != c.subject  # 排除自引用
                )
                if is_intermediate:
                    intermediate_certs.append((filename, cert))
                else:
                    leaf_certs.append((filename, cert))
        
        print(f"根证书: {len(root_certs)}, 中间证书: {len(intermediate_certs)}, 叶子证书: {len(leaf_certs)}")
        
        # 为每个证书构建链
        for filename, cert in self.all_certs:
            chain = self.find_certificate_chain(cert, root_certs, intermediate_certs)
            self.cert_chains[filename] = chain
        
        return self.cert_chains
    
    def find_certificate_chain(self, cert, root_certs, intermediate_certs):
        """查找证书的完整链"""
        chain = [cert]
        current = cert
        
        # 向上查找颁发者，最多10级
        for _ in range(10):
            # 在中间证书中查找颁发者
            issuer_found = None
            for _, potential_issuer in intermediate_certs + root_certs:
                if (potential_issuer.subject == current.issuer and 
                    potential_issuer != current):  # 避免自引用
                    issuer_found = potential_issuer
                    break
            
            if issuer_found:
                chain.append(issuer_found)
                current = issuer_found
                
                # 如果找到根证书，停止
                if issuer_found.issuer == issuer_found.subject:
                    break
            else:
                # 找不到颁发者，链不完整
                break
        
        return chain
    
    def analyze_trust_status(self):
        """分析所有证书的信任状态"""
        print("分析信任状态...")
        
        results = []
        trusted_count = 0
        
        for filename, cert in self.all_certs:
            chain = self.cert_chains.get(filename, [cert])
            chain_length = len(chain)
            
            # 检查链的完整性
            is_chain_complete = False
            trust_status = "未知"
            
            if chain_length == 1:
                # 单个证书
                if cert.issuer == cert.subject:
                    trust_status = "自签名根证书"
                    is_chain_complete = True
                else:
                    trust_status = "证书链不完整（缺少颁发者）"
            else:
                # 检查链是否以根证书结束
                root_cert = chain[-1]
                if root_cert.issuer == root_cert.subject:
                    trust_status = f"完整证书链 ({chain_length} 级)"
                    is_chain_complete = True
                    trusted_count += 1
                else:
                    trust_status = f"不完整证书链 ({chain_length} 级，缺少根证书)"
            
            # 检查有效期
            current_time = datetime.now(timezone.utc)
            is_valid = (cert.not_valid_before_utc <= current_time <= cert.not_valid_after_utc)
            
            results.append({
                'filename': filename,
                'subject': cert.subject.rfc4514_string(),
                'issuer': cert.issuer.rfc4514_string(),
                'chain_length': chain_length,
                'is_chain_complete': is_chain_complete,
                'trust_status': trust_status,
                'is_valid': is_valid,
                'valid_from': cert.not_valid_before_utc,
                'valid_to': cert.not_valid_after_utc
            })
        
        trusted_percentage = (trusted_count / len(results)) * 100 if results else 0
        
        return results, trusted_percentage
    
    def generate_detailed_report(self, results, trusted_percentage, output_file):
        """生成详细报告"""
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write("证书信任分析详细报告\n")
            f.write("=" * 80 + "\n")
            f.write(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"分析证书数量: {len(results)}\n")
            f.write(f"可信证书比例: {trusted_percentage:.1f}%\n")
            f.write("=" * 80 + "\n\n")
            
            # 按信任状态分组
            by_status = collections.defaultdict(list)
            for result in results:
                by_status[result['trust_status']].append(result)
            
            # 输出统计
            f.write("信任状态统计:\n")
            for status, certs in by_status.items():
                f.write(f"- {status}: {len(certs)} 个证书 ({len(certs)/len(results)*100:.1f}%)\n")
            f.write("\n")
            
            # 详细列表
            f.write("证书详细信息:\n")
            f.write("=" * 120 + "\n")
            
            for i, result in enumerate(results, 1):
                status_icon = "✓" if result['is_chain_complete'] else "✗"
                valid_icon = "✓" if result['is_valid'] else "✗"
                
                f.write(f"{i}. {result['filename']}\n")
                f.write(f"   主题: {result['subject']}\n")
                f.write(f"   颁发者: {result['issuer']}\n")
                f.write(f"   证书链: {result['chain_length']} 级 | {status_icon} {result['trust_status']}\n")
                f.write(f"   有效期: {result['valid_from'].strftime('%Y-%m-%d')} 至 {result['valid_to'].strftime('%Y-%m-%d')}\n")
                f.write(f"   状态: {valid_icon} {'有效' if result['is_valid'] else '过期'}\n")
                
                # 显示证书链（如果是多级）
                if result['chain_length'] > 1:
                    f.write("   证书链:\n")
                    chain = self.cert_chains.get(result['filename'], [])
                    for j, chain_cert in enumerate(chain):
                        indent = "      " + "  " * j
                        subject_short = chain_cert.subject.rfc4514_string()[:80] + "..." if len(chain_cert.subject.rfc4514_string()) > 80 else chain_cert.subject.rfc4514_string()
                        f.write(f"{indent}{j+1}. {subject_short}\n")
                
                f.write("-" * 120 + "\n")
        
        print(f"详细报告已保存到: {output_file}")
        print(f"最终可信证书比例: {trusted_percentage:.1f}%")
        
        # 提供改进建议
        if trusted_percentage < 100:
            print("\n改进建议:")
            print("1. 确保所有中间证书都已安装")
            print("2. 检查是否有证书过期")
            print("3. 运行证书缓存刷新工具")
            print("4. 重启计算机应用更改")

def main():
    if len(sys.argv) != 2:
        print("使用方法: python cert_analyzer_pro.py <证书文件夹>")
        sys.exit(1)
    
    folder_path = sys.argv[1]
    if not os.path.isdir(folder_path):
        print(f"错误: 文件夹不存在: {folder_path}")
        sys.exit(1)
    
    # 创建分析器
    analyzer = CertificateAnalyzer(folder_path)
    
    # 加载证书
    analyzer.load_all_certificates()
    
    # 构建证书链
    analyzer.build_certificate_chains()
    
    # 分析信任状态
    results, trusted_percentage = analyzer.analyze_trust_status()
    
    # 生成报告
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"Certificate_Analysis_Report_{timestamp}.txt"
    analyzer.generate_detailed_report(results, trusted_percentage, output_file)

if __name__ == "__main__":
    main()