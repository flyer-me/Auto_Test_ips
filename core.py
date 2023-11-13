# 导入所需的模块
import socket
import threading
import time
import queue

# 定义一个函数，用于测试一个ip地址是否连通
def test_ip(ip, q):
    # 创建一个套接字对象
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    # 设置超时时间为1秒
    s.settimeout(1)
    # 尝试连接ip地址的80端口
    try:
        s.connect((ip, 80))
        # 如果连接成功，说明ip地址连通，关闭套接字并返回
        s.close()
        return
    except:
        # 如果连接失败，说明ip地址不连通，将ip地址放入队列中，并关闭套接字
        q.put(ip)
        s.close()

# 定义一个函数，用于从队列中取出不连通的ip地址，并写入bad.txt文件
def write_bad(q):
    # 打开bad.txt文件，追加模式
    with open("bad.txt", "a") as f:
        # 循环从队列中取出ip地址，直到队列为空
        while not q.empty():
            # 取出一个ip地址
            ip = q.get()
            # 写入文件
            f.write(ip + "\n")
            # 标记任务完成
            q.task_done()

# 定义一个主函数，用于读取ip.txt文件，创建多线程，测试ip地址，计时，重复测试
def main():
    # 读取ip.txt文件，将每一行的ip地址存入一个列表
    with open("ip.txt", "r") as f:
        ips = f.read().splitlines()
    # 定义一个变量，用于记录测试次数
    count = 0
    # 定义一个循环，用于重复测试
    while count < 30:
        # 清空bad.txt文件
        open("bad.txt", "w").close()
        # 记录测试开始的时间
        start = time.time()
        # 创建一个队列对象，用于存储不连通的ip地址
        q = queue.Queue()
        # 创建一个空列表，用于存储线程对象
        threads = []
        # 遍历ip地址列表，为每一个ip地址创建一个线程，调用test_ip函数，将线程对象添加到列表中
        for ip in ips:
            t = threading.Thread(target=test_ip, args=(ip, q))
            t.start()
            threads.append(t)
        # 遍历线程列表，等待每一个线程结束
        for t in threads:
            t.join()
        # 创建一个线程，调用write_bad函数，将不连通的ip地址写入文件
        t = threading.Thread(target=write_bad, args=(q,))
        t.start()
        # 等待队列中的所有任务完成
        q.join()
        # 记录测试结束的时间
        end = time.time()
        # 计算测试所用的时间
        duration = end - start
        # 打印测试结果
        print(f"第{count + 1}次测试，用时{duration}秒，不连通的ip地址已写入bad.txt文件")
        # 增加测试次数
        count += 1

# 调用主函数
if __name__ == "__main__":
    main()
