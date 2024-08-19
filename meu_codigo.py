from abc import ABC




class Investir(ABC):
    def __init__(self) -> None:
        pass

    def creditoInter(self, *args, **kwargs) -> None:
        pass

    def debitoInter(self, *args, **kwargs) -> None:
        pass


class CartaoDjalma(Investir):
    def creditoInter(self, *args, **kwargs) -> None:
        print("Oi")
        return super().creditoInter(*args, **kwargs)



def main(*args, **kwargs) -> None:
    i1 = Investir()

    print(i1)


if __name__ == "__main__":
    main()
