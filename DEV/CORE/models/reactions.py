from dataclasses import dataclass
from typing import Optional

@dataclass
class Reaction:
    """
    Represente une réaction d'appuie 

    Exemple:
    Ay = Reaction verticale en A
    Bx = Reaction horizontale en B
    MA = Moment d'encastrement en A
    """
    nom: str
    x: float
    direction: str
    valeur: Optional[float] = None

    def is_known(self) -> bool:
        return self.valeur is not None
    
