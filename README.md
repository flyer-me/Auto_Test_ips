# auto_test_ips_in_excel
auto test ips in excel(s) and write back.

自动使用ping测试位于Excel文件中的IP地址，并将连通性结果写回Excel文件.
由于项目历史遗留,Excel文件中IP地址表头需为以下中的一个：
['专网地址', '专网IP', '专网ip', '外网IP']
你可以根据需要修改，或者修改代码为从配置文件中读取.

Refs: PingInfoView,https://www.nirsoft.net/utils/multiple_ping_tool.html
