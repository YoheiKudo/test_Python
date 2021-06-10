class Item():
    """
    初期化時に変数名の前に__をつけると外部参照できなくすることができる
    __で始まる変数は同じクラス内であれば自由に使える
    """

    def __init__(self, name, price):
        self.__data = {"name": name, "price": price}

    """
    pythonにおけるgetter
    プロパティ前に@propertyを付ける
    getter名とsetter名とプロパティ名は同じ名前にすること
    """

    @property
    def name(self):
        return self.__data["name"]

    """
    pythonにおけるsetter
    関数前に@(setter名).propertyを付ける
    getter名とsetter名とプロパティ名は同じ名前にすること
    """

    @name.setter
    def name(self, value):
        self.__data["name"] = value

    """
    getterのみ設定するとreadonlyになる
    """

    @property
    def price(self):
        return self.__data["price"]


# サンプル
"""インスタンス化"""
watch = Item("Rolex", 100000)
"""プロパティの参照"""
print(watch.name)
print(watch.price)
"""プロパティの更新"""
watch.name = "Omega"
print(watch.name)
# """priceはgetterのみなので、プロパティの更新は出来ない"""
# watch.price=200000
# print(watch.price)
