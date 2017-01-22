# LikeArrayList
VBA Array WrapperClass Like .NET ArrayList.

.NETのArrayList風のメソッドを持つ、VBAの配列ラッパークラスです。

# 特徴

+ VBAでよく使われるコレクションとの対応をとるため、基本的には添え字は1から始まります。
+ InitInternalArrayメソッドに値型の配列を指定することで、中の配列を指定した値型に変更できます（要素追加時に自動で型変換されます）。
+ InitInternalArrayメソッドにオブジェクト型の配列を指定した場合は、自動でObject型になるため、値型しか除外できません。
