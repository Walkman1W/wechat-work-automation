import pandas as pd

# 创建示例数据
data = {
    '序号': range(1, 4),  # 1到10的序号
    '姓名': ['郭峰', '张三', '李四'],
    'phone': [  # 随机生成的手机号
        '13062812446',
        '13912345678',
        '13812345679'
    ]
}

# 创建DataFrame
df = pd.DataFrame(data)

# 保存到Excel
df.to_excel('contacts.xlsx', index=False)
print("Excel文件已创建完成！") 