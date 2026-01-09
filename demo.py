import can
import time
from can import Message

# 配置CAN接口参数
interface = 'pcan'  # 根据系统调整，如Windows使用'pcan'、Linux使用'socketcan'
channel = 'PCAN_USBBUS1'         # CAN接口名称
bitrate = 250000         # CAN总线速率

# 扫描的CAN ID范围（示例范围：0x700到0x7FF）
start_id = 0x000
end_id = 0x7FF

# UDS请求配置
uds_service = 0x2E       # TesterPresent服务ID
uds_subfunc = 0x01       # 子功能（可选）
request_data = [0x02, uds_service, uds_subfunc] + [0x00] * 5  # 单帧请求：长度+服务+子功能

# 初始化CAN接口
try:
    bus = can.Bus(interface=interface, channel=channel, bitrate=bitrate, receive_own_messages=False)
except Exception as e:
    print(f"初始化CAN接口失败: {e}")
    exit()

def check_uds_response(received_data):
    """检查接收的数据是否符合UDS肯定响应格式"""
    if len(received_data) < 2:
        return False
    # 响应第二个字节应为服务ID + 0x40（例如0x3E -> 0x7E）
    return received_data[1] == (uds_service + 0x40)

def scan_uds_ids():
    valid_ids = []
    print(f"开始扫描CAN ID范围: 0x{start_id:03X} 到 0x{end_id:03X}...")

    for can_id in range(start_id, end_id + 1):
        # 构造CAN消息
        msg = Message(
            arbitration_id=can_id,
            data=request_data,
            is_extended_id=False  # 标准11位ID
        )

        try:
            # 发送请求
            bus.send(msg)
            time.sleep(0.05)  # 等待响应

            # 接收并检查响应
            response = bus.recv(timeout=0.1)
            print(response)
            if response and check_uds_response(response.data):
                valid_ids.append(can_id)
                print(f"发现UDS节点: 0x{can_id:03X} 响应来自ID: 0x{response.arbitration_id:03X}")

        except can.CanError as e:
            print(f"发送/接收错误: {e}")

    return valid_ids

if __name__ == "__main__":
    try:
        detected_ids = scan_uds_ids()
        print("\n扫描完成。检测到以下ID支持UDS:")
        for uid in detected_ids:
            print(f"0x{uid:03X}")
    except KeyboardInterrupt:
        print("\n用户中断扫描。")
    finally:
        bus.shutdown()