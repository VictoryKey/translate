import os
import sys
import json
from pathlib import Path
from datetime import datetime

class MultiLevelFolderCreator:
    def __init__(self):
        self.config_file = "folder_structure_config.json"
        self.templates_file = "folder_templates.json"
        
    def display_menu(self):
        """显示主菜单"""
        print("\n" + "="*60)
        print("多级文件夹批量创建工具")
        print("="*60)
        print("1. 创建新的文件夹结构")
        print("2. 从模板创建")
        print("3. 保存当前配置为模板")
        print("4. 管理模板")
        print("5. 批量创建多个结构")
        print("6. 退出")
        print("="*60)
        
    def get_user_input(self, prompt, input_type=str, default=None, options=None):
        """获取用户输入，支持默认值和选项验证"""
        while True:
            try:
                if default is not None:
                    prompt += f" (默认: {default}): "
                else:
                    prompt += ": "
                    
                user_input = input(prompt).strip()
                
                if not user_input and default is not None:
                    return default
                    
                if input_type == int:
                    result = int(user_input)
                    if options and result not in options:
                        print(f"请输入有效的选项: {options}")
                        continue
                    return result
                elif input_type == bool:
                    if user_input.lower() in ['y', 'yes', '是', '1']:
                        return True
                    elif user_input.lower() in ['n', 'no', '否', '0']:
                        return False
                    else:
                        print("请输入 y/是 或 n/否")
                        continue
                else:
                    return user_input
                    
            except ValueError:
                print("输入格式错误，请重新输入")
            except KeyboardInterrupt:
                print("\n操作已取消")
                return None
    
    def create_folder_structure_interactive(self):
        """交互式创建文件夹结构"""
        print("\n开始创建文件夹结构...")
        
        # 获取根目录
        root_path = self.get_user_input("请输入根目录路径", default=".")
        if root_path is None:
            return
            
        root_path = Path(root_path).resolve()
        
        # 创建根目录（如果不存在）
        root_path.mkdir(parents=True, exist_ok=True)
        print(f"根目录: {root_path}")
        
        # 获取层级数
        max_levels = self.get_user_input("请输入要创建的文件夹层级数", int, default=3)
        if max_levels is None:
            return
            
        # 存储结构配置
        structure = {
            "root": str(root_path),
            "levels": []
        }
        
        # 为每一级配置
        for level in range(1, max_levels + 1):
            print(f"\n--- 配置第 {level} 级文件夹 ---")
            
            # 获取文件夹数量
            num_folders = self.get_user_input(f"第{level}级要创建几个文件夹", int, default=3)
            if num_folders is None:
                return
                
            # 获取命名模式
            print(f"\n第{level}级文件夹命名选项:")
            print("1. 使用固定前缀+数字 (如: Folder_01, Folder_02)")
            print("2. 使用自定义名称列表")
            print("3. 使用字母序列 (如: A, B, C)")
            print("4. 使用日期格式")
            print("5. 使用层级标记")
            
            naming_choice = self.get_user_input("请选择命名模式", int, default=1, options=[1, 2, 3, 4, 5])
            if naming_choice is None:
                return
                
            level_config = {
                "level": level,
                "num_folders": num_folders,
                "naming_mode": naming_choice
            }
            
            if naming_choice == 1:
                # 固定前缀+数字
                prefix = self.get_user_input("请输入文件夹名前缀", default=f"Level{level}")
                if prefix is None:
                    return
                level_config["prefix"] = prefix
                level_config["start_number"] = self.get_user_input("起始数字", int, default=1)
                level_config["zero_pad"] = self.get_user_input("是否用0填充数字 (y/n)", bool, default=True)
                
            elif naming_choice == 2:
                # 自定义名称列表
                print(f"请输入第{level}级{num_folders}个文件夹的名称，用逗号分隔:")
                names_input = self.get_user_input("文件夹名称")
                if names_input is None:
                    return
                names = [name.strip() for name in names_input.split(",")]
                if len(names) != num_folders:
                    print(f"错误: 输入了{len(names)}个名称，但需要{num_folders}个")
                    # 自动补全或截断
                    if len(names) < num_folders:
                        for i in range(len(names), num_folders):
                            names.append(f"Folder_{i+1}")
                    else:
                        names = names[:num_folders]
                level_config["custom_names"] = names
                
            elif naming_choice == 3:
                # 字母序列
                start_char = self.get_user_input("起始字母", default="A")
                if start_char is None:
                    return
                level_config["start_char"] = start_char.upper()
                
            elif naming_choice == 4:
                # 日期格式
                date_format = self.get_user_input("日期格式", default="%Y%m%d")
                if date_format is None:
                    return
                level_config["date_format"] = date_format
                level_config["date_offset"] = self.get_user_input("日期偏移量(天)", int, default=0)
                
            elif naming_choice == 5:
                # 层级标记
                level_config["include_parent"] = self.get_user_input("是否包含父级名称", bool, default=True)
                
            structure["levels"].append(level_config)
            
            # 询问是否继续下一级
            if level < max_levels:
                continue_next = self.get_user_input("是否继续配置下一级", bool, default=True)
                if not continue_next:
                    break
        
        return structure
    
    def generate_folder_names(self, level_config, parent_path="", level_path=""):
        """根据配置生成文件夹名称"""
        num_folders = level_config["num_folders"]
        naming_mode = level_config["naming_mode"]
        folder_names = []
        
        if naming_mode == 1:
            # 固定前缀+数字
            prefix = level_config.get("prefix", f"Level{level_config['level']}")
            start_num = level_config.get("start_number", 1)
            zero_pad = level_config.get("zero_pad", True)
            
            for i in range(num_folders):
                num = start_num + i
                if zero_pad:
                    # 根据数量决定填充位数
                    digits = len(str(num_folders + start_num - 1))
                    folder_name = f"{prefix}_{num:0{digits}d}"
                else:
                    folder_name = f"{prefix}_{num}"
                folder_names.append(folder_name)
                
        elif naming_mode == 2:
            # 自定义名称
            custom_names = level_config.get("custom_names", [])
            if len(custom_names) >= num_folders:
                folder_names = custom_names[:num_folders]
            else:
                # 如果自定义名称不够，用默认名称补全
                folder_names = custom_names
                for i in range(len(custom_names), num_folders):
                    folder_names.append(f"Folder_{i+1}")
                    
        elif naming_mode == 3:
            # 字母序列
            start_char = level_config.get("start_char", "A")
            start_code = ord(start_char.upper())
            
            for i in range(num_folders):
                char_code = start_code + i
                if char_code <= 90:  # Z的ASCII码
                    folder_name = chr(char_code)
                else:
                    # 超过Z，使用AA, AB, AC...
                    first_letter = chr(start_code + (i // 26) - 1)
                    second_letter = chr(start_code + (i % 26))
                    folder_name = first_letter + second_letter
                folder_names.append(folder_name)
                
        elif naming_mode == 4:
            # 日期格式
            date_format = level_config.get("date_format", "%Y%m%d")
            date_offset = level_config.get("date_offset", 0)
            base_date = datetime.now()
            
            for i in range(num_folders):
                from datetime import timedelta
                current_date = base_date + timedelta(days=date_offset + i)
                folder_name = current_date.strftime(date_format)
                folder_names.append(folder_name)
                
        elif naming_mode == 5:
            # 层级标记
            include_parent = level_config.get("include_parent", True)
            parent_name = Path(parent_path).name if include_parent and parent_path else ""
            
            for i in range(num_folders):
                if parent_name:
                    folder_name = f"{parent_name}_L{level_config['level']}_{i+1}"
                else:
                    folder_name = f"L{level_config['level']}_{i+1}"
                folder_names.append(folder_name)
        
        return folder_names
    
    def create_folders(self, structure, current_path="", level_index=0):
        """递归创建文件夹结构"""
        if level_index >= len(structure["levels"]):
            return 1  # 叶节点计数
        
        level_config = structure["levels"][level_index]
        folder_names = self.generate_folder_names(level_config, current_path, f"L{level_index+1}")
        
        total_created = 0
        
        for folder_name in folder_names:
            folder_path = Path(current_path) / folder_name
            folder_path.mkdir(parents=True, exist_ok=True)
            print(f"创建: {folder_path}")
            
            # 递归创建子文件夹
            sub_created = self.create_folders(structure, folder_path, level_index + 1)
            total_created += sub_created
        
        return total_created * len(folder_names) if level_index > 0 else total_created
    
    def save_template(self, structure, template_name):
        """保存配置为模板"""
        try:
            # 加载现有模板
            templates = self.load_templates()
            
            # 添加新模板
            templates[template_name] = {
                "structure": structure,
                "created_date": datetime.now().isoformat(),
                "description": structure.get("description", "")
            }
            
            # 保存模板
            with open(self.templates_file, 'w', encoding='utf-8') as f:
                json.dump(templates, f, indent=2, ensure_ascii=False)
                
            print(f"✓ 模板 '{template_name}' 已保存")
            return True
            
        except Exception as e:
            print(f"保存模板失败: {e}")
            return False
    
    def load_templates(self):
        """加载模板"""
        try:
            if os.path.exists(self.templates_file):
                with open(self.templates_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return {}
        except:
            return {}
    
    def list_templates(self):
        """列出所有模板"""
        templates = self.load_templates()
        if not templates:
            print("暂无模板")
            return []
            
        print("\n可用模板:")
        print("-" * 50)
        for i, (name, template) in enumerate(templates.items(), 1):
            desc = template.get("description", "无描述")
            date = template.get("created_date", "未知日期")
            print(f"{i}. {name} - {desc} ({date})")
        print("-" * 50)
        
        return list(templates.keys())
    
    def create_from_template(self):
        """从模板创建文件夹结构"""
        templates = self.load_templates()
        if not templates:
            print("暂无可用模板")
            return None
            
        template_names = self.list_templates()
        choice = self.get_user_input("请选择模板编号", int, options=range(1, len(template_names) + 1))
        if choice is None:
            return None
            
        template_name = template_names[choice - 1]
        structure = templates[template_name]["structure"]
        
        # 允许修改根目录
        new_root = self.get_user_input("输入新的根目录（留空使用模板配置）", default="")
        if new_root:
            structure["root"] = new_root
            
        return structure
    
    def batch_create_structures(self):
        """批量创建多个文件夹结构"""
        print("\n批量创建多个文件夹结构")
        
        num_structures = self.get_user_input("要创建几个不同的文件夹结构", int, default=3)
        if num_structures is None:
            return
            
        base_root = self.get_user_input("基础根目录", default=".")
        if base_root is None:
            return
            
        structures = []
        
        for i in range(num_structures):
            print(f"\n--- 配置第 {i+1} 个结构 ---")
            structure_root = self.get_user_input(f"结构{i+1}的根目录", default=f"{base_root}/Structure_{i+1}")
            if structure_root is None:
                return
                
            # 简化配置：只问层级数和每级文件夹数
            levels = self.get_user_input("层级数", int, default=3)
            if levels is None:
                return
                
            structure = {"root": structure_root, "levels": []}
            
            for level in range(1, levels + 1):
                num_folders = self.get_user_input(f"第{level}级文件夹数", int, default=3)
                if num_folders is None:
                    return
                    
                structure["levels"].append({
                    "level": level,
                    "num_folders": num_folders,
                    "naming_mode": 1,  # 默认使用前缀+数字
                    "prefix": f"L{level}",
                    "start_number": 1,
                    "zero_pad": True
                })
                
            structures.append(structure)
            
        # 批量创建
        total_created = 0
        for i, structure in enumerate(structures, 1):
            print(f"\n创建结构 {i}/{len(structures)}: {structure['root']}")
            created = self.create_folders(structure, structure["root"], 0)
            total_created += created
            print(f"结构 {i} 创建完成")
            
        print(f"\n批量创建完成! 总共创建了 {total_created} 个文件夹")
        return total_created
    
    def manage_templates(self):
        """管理模板"""
        while True:
            print("\n模板管理")
            print("1. 查看模板列表")
            print("2. 删除模板")
            print("3. 返回主菜单")
            
            choice = self.get_user_input("请选择", int, options=[1, 2, 3])
            if choice is None:
                break
                
            if choice == 1:
                self.list_templates()
            elif choice == 2:
                self.delete_template()
            elif choice == 3:
                break
    
    def delete_template(self):
        """删除模板"""
        template_names = self.list_templates()
        if not template_names:
            return
            
        choice = self.get_user_input("选择要删除的模板编号", int, options=range(1, len(template_names) + 1))
        if choice is None:
            return
            
        template_name = template_names[choice - 1]
        confirm = self.get_user_input(f"确认删除模板 '{template_name}'?", bool, default=False)
        
        if confirm:
            templates = self.load_templates()
            if template_name in templates:
                del templates[template_name]
                with open(self.templates_file, 'w', encoding='utf-8') as f:
                    json.dump(templates, f, indent=2, ensure_ascii=False)
                print(f"✓ 模板 '{template_name}' 已删除")
    
    def run(self):
        """运行主程序"""
        print("多级文件夹批量创建工具 v1.0")
        
        while True:
            self.display_menu()
            choice = self.get_user_input("请选择操作", int, options=[1, 2, 3, 4, 5, 6])
            
            if choice is None:
                continue
                
            if choice == 1:
                # 创建新的文件夹结构
                structure = self.create_folder_structure_interactive()
                if structure:
                    print(f"\n开始创建文件夹结构...")
                    total_created = self.create_folders(structure, structure["root"], 0)
                    print(f"\n✓ 创建完成! 总共创建了 {total_created} 个文件夹")
                    
                    # 询问是否保存为模板
                    save_template = self.get_user_input("是否保存此配置为模板", bool, default=False)
                    if save_template:
                        template_name = self.get_user_input("输入模板名称")
                        if template_name:
                            structure["description"] = self.get_user_input("模板描述", default="")
                            self.save_template(structure, template_name)
                
            elif choice == 2:
                # 从模板创建
                structure = self.create_from_template()
                if structure:
                    print(f"\n开始从模板创建文件夹...")
                    total_created = self.create_folders(structure, structure["root"], 0)
                    print(f"\n✓ 创建完成! 总共创建了 {total_created} 个文件夹")
                
            elif choice == 3:
                # 保存当前配置为模板
                structure = self.create_folder_structure_interactive()
                if structure:
                    template_name = self.get_user_input("输入模板名称")
                    if template_name:
                        structure["description"] = self.get_user_input("模板描述", default="")
                        self.save_template(structure, template_name)
                
            elif choice == 4:
                # 管理模板
                self.manage_templates()
                
            elif choice == 5:
                # 批量创建多个结构
                self.batch_create_structures()
                
            elif choice == 6:
                print("感谢使用! 再见!")
                break
                
            # 暂停一下让用户看到结果
            if choice != 6:
                input("\n按Enter键继续...")

def main():
    """主函数"""
    try:
        creator = MultiLevelFolderCreator()
        creator.run()
    except KeyboardInterrupt:
        print("\n\n程序被用户中断")
    except Exception as e:
        print(f"程序执行出错: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()