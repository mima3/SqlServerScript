SqlServerScript
==========
このスクリプトはSQLServerの操作を行うためのVBScriptです。

GetSpList
------
ストアドプロシージャの一覧を取得します。

### 実行例 ###
    CScript GetSpList.wsf ServerName DBName UserName Pass

GetSpMd5
------
ストアドプロシージャの一覧を取得してMD5値に変更します。
これを利用して、複数のDBでストアドの変更の有無をチェックすることができます。 

### 実行例 ###
    CScript GetSpMd5.wsf ServerName DBName UserName Pass

DependSp
------
ストアドプロシージャの依存しているオブジェクトを取得します

### 実行例 ###
    CScript DependSp.wsf ServerName DBName UserName Pass

