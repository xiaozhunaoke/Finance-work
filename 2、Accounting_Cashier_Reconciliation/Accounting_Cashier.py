import os
import pandas as pd
"""1、读取会计账、出纳账分别存储在DataFrme容器中；
2、按照dataFrame1的数据，遍历dataFrame2的数据，如果找到第一个相同的，在原数据上同时删除两者
3、剩下的就为不同数据，写入新的Excel中"""

def xlfs_path(xlfs):#返回文件的路径为当前文件夹的路径
    return os.path.abspath('.')+'\\'+xlfs

def read_data(xlfs,col):#读取特定列的数据
    path=xlfs_path(xlfs)
    data=pd.read_excel(path,usecols=str(col))#读取列名为col列的数据
    data=data.iloc[7:]#切片，删除不需要的数据
    data=data.dropna().reset_index(drop=True)#删除缺失值，然后进行索引重新排列
    data.columns=['摘要','数据']
    data=data[~ data['摘要'].str.contains("期初|合计|累计|小计",na=False)].reset_index(drop=True)#去除期初、合计、小计累计等数
    return data

def diff_data(ac_data,ca_data):#依次按照dataFrame1的数据，遍历dataFrame2的数据，如果找到第一个相同的，同时删除两者，结束该次循环，继续dataFrame1中下一个数据的寻找
    data=pd.DataFrame()
    l=len(ac_data)
    for i in range(l):#按会计账的数据遍历出纳账的数据
        m=len(ca_data)
        for j in range(m):
            if abs(float(ac_data['数据'].loc[i])-float(ca_data['数据'].loc[j]))<(1e-5):#删除数据后隐式索引重新排列，但显示索引仍保留原来的，所以需用显示索引          
                ac_data.drop([i],inplace=True)#inplace在原数据上删除某行，会改变原数据
                ca_data.drop([j],inplace=True)#drop函数不会删除原显示索引
                ca_data.reset_index(drop=True,inplace=True)
                break
    ac_data.reset_index(drop=True,inplace=True)
    return ac_data,ca_data

def wirte_data():
    xlfs=['会计账.xls','出纳账.xls']
    col=['D,F','A,K','D,G','A,L']
    name=['会计收入差异','出纳收入差异','会计支出差异','出纳支出差异']
    data=pd.DataFrame()
    for n in range(0,3,2):#循环两次操作，分别是收入和支出，分别对比
        ac_data,ca_data=pd.DataFrame(),pd.DataFrame()#在对比完一次时，清空数据容器，进行下一次
        ac_data=read_data(xlfs[0],col[n])#读取会计账的数据
        ca_data=read_data(xlfs[1],col[n+1])#读取出纳账的数据
        ac_data,ca_data=diff_data(ac_data,ca_data)#获取两组数据不同的数
        ac_data.columns=['摘要',name[n]]#对各数据的列标签，重新命名，这样在输出时便于理解
        ca_data.columns=['银行账户',name[n+1]]
        data=pd.concat([data,ac_data,ca_data],axis=1)
    data.to_excel('核对表.xlsx')

if __name__ == "__main__":
    wirte_data()