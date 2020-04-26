import pandas as pd
import os
"""功能：银行数据和出纳数据进行核对：
1、先取多家银行的可以数据取出来，关键行：银行账号，日期，借方金额，贷方金额，每家银行的格式不一致，则每家银行单独读取
2、银行数据较多，单独放在‘银行流水’文件夹名中
3、读取出纳数据
4、整理后的银行数据与出纳数据进行核对，逐项删除相同项
"""

#读取银行数据，因为每家银行的格式不一致，每家银行的读取方式就写一个函数
def read_banks():
    xlfs=read_xlfs()
    bank_data=pd.DataFrame(columns=['账号','交易日期','借方','贷方'],dtype=object)#新建一个空的数据格式
    for i in range(len(xlfs)):
        if ('中国银行' in xlfs[i]) or ('中行' in xlfs[i]):
            data=read_ZGYH(xlfs[i])
            bank_data=pd.concat([bank_data,data])
        elif ('建设银行' in xlfs[i]) or ('建行' in xlfs[i]):
            data=read_JSYH(xlfs[i])
            bank_data=pd.concat([bank_data,data])
        elif ('富滇' in xlfs[i]):
            data=read_FDYH(xlfs[i])
            bank_data=pd.concat([bank_data,data])
        elif ('厦门' in xlfs[i]):
            data=read_XMYH(xlfs[i])
            bank_data=pd.concat([bank_data,data])
        elif ('兴业' in xlfs[i]):
            data=read_XYYH(xlfs[i])
            bank_data=pd.concat([bank_data,data])
        elif ('招商银行' in xlfs[i]) or ('招行' in xlfs[i]):
            data=read_ZSYH(xlfs[i])
            bank_data=pd.concat([bank_data,data])
        elif ('中信' in xlfs[i]):
            data=read_ZXYH(xlfs[i])
            bank_data=pd.concat([bank_data,data])
        elif ('民生' in xlfs[i]):
            data=read_MSYH(xlfs[i])
            bank_data=pd.concat([bank_data,data])
        elif ('浦发' in xlfs[i]):
            data=read_PFYH(xlfs[i])
            bank_data=pd.concat([bank_data,data])
        elif ('广发' in xlfs[i]):
            data=read_GFYH(xlfs[i])
            bank_data=pd.concat([bank_data,data])
    bank_data['账号']=bank_data['账号'].astype(str)#整理数据，相同的格式
    bank_data['交易日期']=bank_data['交易日期'].astype(str)
    bank_data['借方']=bank_data['借方'].astype(str).str.replace(',','')
    bank_data['借方']=bank_data['借方'].str.replace('-','0').astype(float)
    bank_data['贷方']=bank_data['贷方'].astype(str).str.replace(',','')
    bank_data['贷方']=bank_data['贷方'].str.replace('-','0').astype(float)
    bank_data['交易日期']=bank_data['交易日期'].str.replace('-','')
    bank_data.to_excel('银行数据.xlsx')
    return bank_data

#银行流水文件夹存放的路径，获得所有的文件夹名称
def read_xlfs():
    path=os.path.abspath('.')+'\\银行流水'
    xlfs = [x for x in os.listdir(path)]
    xlfs=pd.Series(xlfs)
    xlfs=path+'\\'+xlfs
    return xlfs

def read_cashier():#读取出纳账，对需要几项进行筛选
    cashier=pd.read_excel('出纳账.xls')
    index_name=cashier.iloc[6]
    cashier=cashier.rename(columns=index_name)
    cashier=cashier[7:]
    cashier=cashier.iloc[:,[0,4,10,11]].rename(columns={'资金帐户':'账号','单据日期':'交易日期','本币':'借方','支出':'贷方'})
    cashier=cashier[~ cashier['账号'].str.contains("小计|累计|单位",na=False)]#删除包含关键词‘小计、累计、单位’的行
    cashier=cashier[:-1]
    cashier['交易日期']=cashier['交易日期'].str.replace('-','')
    #cashier.set_index(['账号'],inplace=True)
    cashier['账号']=cashier['账号'].astype(str)
    cashier['交易日期']=cashier['交易日期'].astype(str)
    cashier.to_excel('出纳数据.xlsx')
    return cashier

#1、读取中国银行数据
def read_ZGYH(xlfs):
    zh=pd.read_excel(xlfs)
    zh_account=zh.iloc[0,1]
    index_name=zh.iloc[7]
    zh=zh.rename(columns=index_name)
    zh=zh.rename(columns={'收入':'借方','支出':'贷方','交易日期[ Transaction Date ]':'交易日期','业务类型[ Business type ]':'业务类型','交易金额[ Trade Amount ]':'交易金额','交易类型[ Transaction Type ]':'交易类型'})
    zh=zh.iloc[8:,[0,1,10,13]]
    zh=zh[~ zh['业务类型'].str.contains("自动归集|自动下拨",na=False)]
    zh['账号']=zh_account
    zh_debit=zh[zh['交易类型']=='来账'].rename(columns={'交易金额':'借方'})
    zh_debit=zh_debit[['交易日期','账号','借方']]
    zh_lender=zh[zh['交易类型']=='往账'].rename(columns={'交易金额':'贷方'})
    zh_lender['贷方']=-zh_lender['贷方']
    zh_lender=zh_lender[['交易日期','账号','贷方']]
    zh_data=pd.concat([zh_debit,zh_lender])
    return zh_data

#2、读取建设银行数据
def read_JSYH(xlfs):
    js=pd.read_excel(xlfs)
    js=js.rename(columns={'借方发生额（支取）':'贷方','贷方发生额（收入）':'借方','交易时间':'交易日期'})
    js=js[~ js['摘要'].str.contains('资金归集',na=False)]
    js=js.iloc[:,[0,2,3,4]]
    if js.size!=0:
        js['交易日期']=js['交易日期'].str.split(expand = True)
    return js

#3、富滇
def read_FDYH(xlfs):
    fd_yh=pd.read_excel(xlfs,usecols=['账号','交易日期','转入金额','转出金额'])
    fd_yh=fd_yh.rename(columns={'转入金额':'借方','转出金额':'贷方'})
    fd_yh['账号']=fd_yh['账号'].astype(str)
    return fd_yh

#4、厦门
def read_XMYH(xlfs):
    xm=pd.read_excel(xlfs,usecols=['账户账号','交易日期','转出','转入'])
    xm=xm.rename(columns={'转入':'借方','转出':'贷方','账户账号':'账号'})
    xm['交易日期']=xm['交易日期'].str.split(expand = True)
    xm['账号']=xm['账号'].astype(str)
    return xm 

#5、兴业
def read_XYYH(xlfs):
    xy=pd.read_excel(xlfs,usecols=['账号','交易日期','借方金额','贷方金额'])
    xy=xy.rename(columns={'借方金额':'贷方','贷方金额':'借方'})
    xy['交易日期']=xy['交易日期'].str.split(expand = True)
    xy['账号']=xy['账号'].astype(str)
    return xy

#6、招商
def read_ZSYH(xlfs):
    zs=pd.read_excel(xlfs,index=None)
    zs_zh=zs.iloc[0,7]
    index_name=zs.iloc[7]
    zs=zs.rename(columns=index_name)
    zs=zs.rename(columns={'借方金额':'贷方','贷方金额':'借方','交易日':'交易日期'})
    zs=zs[~ zs['交易类型'].str.contains('协议转账',na=False)]
    zs=zs.iloc[8:,[0,4,5]]
    zs['账号']=zs_zh
    return zs

#7、中信银行
def read_ZXYH(xlfs):
    zx=pd.read_excel(xlfs,index=None,na_filter=False)
    zx_account=zx.iloc[4,2]
    index_name=zx.iloc[9]
    zx=zx.rename(columns=index_name)
    zx=zx.rename(columns={'收款发生额':'借方','付款发生额':'贷方'})
    zx=zx.iloc[10:,[0,6,7]]
    zx['账号']=zx_account
    return zx

#8、民生银行
def read_MSYH(xlfs):
    ms=pd.read_excel(xlfs,index=None)
    ms_account=ms.iloc[0,1]
    index_name=ms.iloc[12]
    ms=ms.rename(columns=index_name)
    ms=ms.rename(columns={'借方发生额':'贷方','贷方发生额':'借方'})
    ms=ms.iloc[13:,[0,2,3]]
    ms['账号']=ms_account
    return ms

#9、浦发银行
def read_PFYH(xlfs):
    pf=pd.read_excel(xlfs,index_col=None,header=None)
    pf_account=pf.iloc[0,1]
    index_name=pf.iloc[3]
    pf=pf.rename(columns=index_name)
    pf=pf.rename(columns={'借方金额':'贷方','贷方金额':'借方'})
    pf=pf.iloc[4:,[0,4,5]]
    pf['账号']=pf_account
    return pf

#10、广发银行
def read_GFYH(xlfs):
    gf=pd.read_excel(xlfs)
    gf_account=gf.iloc[0,2]
    index_name=gf.iloc[6]
    gf=gf.rename(columns=index_name)
    gf=gf.rename(columns={'收入':'借方','支出':'贷方','交易时间':'交易日期'})
    gf['交易日期']=gf['交易日期'].str.split(expand = True)
    gf=gf.iloc[7:,[1,2,3]]
    gf['账号']=gf_account
    return gf

def diff_data(bank_diff,company_diff):#依次按照dataFrame1的数据，遍历dataFrame2的数据，如果找到第一个相同的，同时删除两者，结束该次循环，继续dataFrame1中下一个数据的寻找
    data=pd.DataFrame()
    n=len(bank_diff)
    for i in range(n):#按会计账的数据遍历出纳账的数据
        m=len(company_diff)
        for j in range(m):
            bool_values=(bank_diff.loc[i][0]==company_diff.loc[j][0]) and (bank_diff.loc[i][1]==company_diff.loc[j][1]) and (abs(float(bank_diff.loc[i][2])-float(company_diff.loc[j][2]))<(1e-5))
            if bool_values:#删除数据后隐式索引重新排列，但显示索引仍保留原来的，所以需用显示索引 
                bank_diff.drop([i],inplace=True)#inplace在原数据上删除某行，会改变原数据
                company_diff.drop([j],inplace=True)#drop函数不会删除原显示索引         
                company_diff.reset_index(drop=True,inplace=True)
                break
    bank_diff.reset_index(drop=True,inplace=True)
    return bank_diff,company_diff

def wirte_data():
    bank_data=read_banks()
    cashier_data=read_cashier()
    name=['银行借方差异','银行贷方差异','企业借方差异','企业贷方差异']
    revenue=['借方','贷方']
    data=pd.DataFrame()
    for n in range(2):#循环两次操作，分别是收入和支出，分别对比
        bank_diff,company_diff=pd.DataFrame(),pd.DataFrame()#在对比完一次时，清空数据容器，进行下一次
        bank_diff=bank_data[['账号','交易日期',revenue[n]]]
        bank_diff=bank_diff.dropna().reset_index(drop=True)
        bank_diff=bank_diff[bank_diff[revenue[n]]>0.001].reset_index(drop=True)

        company_diff=cashier_data[['账号','交易日期',revenue[n]]]
        company_diff=company_diff.dropna().reset_index(drop=True)
        company_diff=company_diff[company_diff[revenue[n]]>0.001].reset_index(drop=True)

        bank_diff,company_diff=diff_data(bank_diff,company_diff)

        bank_diff=bank_diff.rename(columns={revenue[n]:name[n]})
        company_diff=company_diff.rename(columns={revenue[n]:name[n+2]})
        
        data=pd.concat([data,bank_diff,company_diff],axis=1)
    data.to_excel('核对表.xlsx')

if __name__ == "__main__":
    wirte_data()