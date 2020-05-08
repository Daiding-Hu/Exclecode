import pymssql
# 连接数据库
connect = pymssql.connect(
    host='localhost',
    port=3306,
    user='root',
    passwd='hu12580WEI',
    db='python',
    charset='utf8'
)

# 获取游标
cursor = connect.cursor()

# 插入数据
sql = "INSERT INTO trade (name, account, saving) VALUES ( '%s', '%s', %.2f )"
data = ('雷军', '13512345678', 10000)
cursor.execute(sql % data)
connect.commit()
print('成功插入', cursor.rowcount, '条数据')
