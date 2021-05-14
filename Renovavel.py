from Usina import Usina;
from RecebeDados import RecebeDados;

class Renovavel (Usina):
    
    def __init__(self, recebe_dados, abaRenov, offset, iRenov):
        
        # define fonte_dados como o objeto da classe RecebeDados e o nome da aba com as usinas UHE
        self.nomeAba = abaRenov;
        self.fonte_dados = recebe_dados;
        self.indexUsinaInterno = iRenov;
        
        # a variavel offset e importante pq a quantidade de linhas que devem ser puladas na planilha
        # pode ser diferente do index da usina. 
        self.linhaOffset = offset;
        
        # metodo referente a classe pai
        super(Renovavel, self).__init__(self.fonte_dados, self.nomeAba, self.linhaOffset);
        
        return;