import string

def letras(num: int) -> list[str]:
    if not isinstance(num, int):
        raise TypeError("O parâmetro deve ser um número.")
    
    if num < 0:
        raise ValueError("O parâmetro deve ser um número inteiro positivo.")
    
    if num > 26:
        raise ValueError("O parâmetro não pode ser maior que 26, pois só há 26 letras no alfabeto.")
    
    return [string.ascii_uppercase[i] for i in range(num)]


print(letras(4))