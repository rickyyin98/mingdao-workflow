# mingdao-workflow  

## 说明

从 [句子互动商业化管理系统-成交](https://www.mingdao.com/app/d497b33d-96a0-4d57-a2fb-307740fd27ba/60d8851fc4418130d0fcef0d/60d8854cf6225cbaa249fa34/60d8854cf6225cbaa249fa38)中的生数据，生成可管理的权责发生制的 Excel。

## 输入：

|  客户      | 行业  | 客户类型  | SKU      | 开始时间 | 持续月份(个月) | 购买数量	| 购买单价  |
| --------  | ---- | -------- | -------- |-------  | ------------ | -------- |------   |
| XXXX      | XXXX | XXXX     |  XXXX    |  XXXX   |  XXXX        |  XXXX    | XXXX    |

## 输出

Excel 包含以下5个sheet：

### 以客户数量为单位的 sheet：
- MRR: 按照权责发生制计算的 MRR，账期为的环境变量为 ACCOUNT_PERIOD， 详细说明见： [MRR 增加账期](https://github.com/juzibot/mingdao-workflow/issues/5)
- raw-MRR: 没有计算账期情况下的 MRR

|  客户名称  | 行业  | 客户类型  | SKU      | 2020-7  | ...    | 2021-12  |
| --------  | ---- | -------- | -------- |------- | ------ | -------- |
| XXXX      | XXXX | XXXX     |  XXXX    |  XXXX  |  XXXX  |  XXXX    |

### 以合同数量为单位的 sheet：
- revenue = price * count
- price = 当前这笔合同的客单价
- count = 当前这笔合同的购买数量

|  客户名称  | 行业  | 客户类型  | SKU      | 2020-7  | ...    | 2021-12  |
| --------  | ---- | -------- | -------- |------- | ------ | -------- |
| XXXX      | XXXX | XXXX     |  XXXX    |  XXXX  |  XXXX  |  XXXX    |


### 类别说明

#### SKU

- 句子秒回
- Wechaty

以下部分没有计入
- 私有部署
- 咨询培训

#### 客户类型

- KA
- SMB
- 代理商

#### 行业

@伯昊 补充


## Installation  

`npm i`  

## Run  

`node index.js yourexcel.xlsx`  

