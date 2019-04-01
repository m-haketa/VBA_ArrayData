# Default Propertyで挙動が変わる?

クラスモジュールのDefault Propertyを使うかどうかで挙動が変わっているように見えます。


## Sub 動く()

最後の行で、AD.Dataの返り値が2次元配列なので、A2セル～C2セルに値が出力されます。



## Sub 動かない()

Default PropertyでDataを指定しているので、先ほどのものと同じ挙動になるはず。

でも、実際には、うまく動きません。

