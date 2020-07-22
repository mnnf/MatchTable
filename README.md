# MatchTable



## 概要

囲碁大会用対局者リストのエクセルをpythonで読み書きすることで対局の組み合わせと成績計算を行う。



## 使用方法

- エクセル上に対局者と棋力を入力し、対局回数に応じて横幅を調整する。
- CI版ではMatchTable.pyに引数を与えてエクセルを更新する。
- UI版ではMatchTableUI.pyを起動する。



MatchTable.pyのコマンド説明

```shell
$ python MatchTable.py エクセルファイル名 コマンド
```

コマンドは'result'で成績の書き込み。数値で対局回数の指定を行う。





(例)第１回戦の組み合わせを作成する。

```shell
$ python MatchTable.py 対局者一覧.xlsx 1
```



(例)第２回戦の組み合わせを作成する。

```shell
$ python MatchTable.py 対局者一覧.xlsx 2
```



(例)成績を作成する。

```shell
$ python MatchTable.py 対局者一覧.xlsx result
```
