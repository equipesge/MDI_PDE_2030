from Usina import Usina;
from RecebeDados import RecebeDados;

class Termica(Usina):
    
    def __init__(self, recebe_dados, abaTerm, offset, iTerm, nMeses, nMesesPos, continuidade=False):
        
        # define fonte_dados como o objeto da classe RecebeDados e o nome da aba com as usinas UHE
        self.nomeAba = abaTerm;
        self.fonte_dados = recebe_dados;
        self.indexUsinaInterno = iTerm;
        self.numMeses = int(nMeses);
        self.numMesesPos = int(nMesesPos);
        
        # a variavel offset e importante pq a quantidade de linhas que devem ser puladas na planilha
        self.linhaOffset = offset;
        
        # metodo referente a classe pai
        super(Termica, self).__init__(self.fonte_dados, self.nomeAba, self.linhaOffset);
        
        # importa dados especificos de UHEs
        if not(hasattr(self,"isProjeto")):
            self.importaDadosTermica();
        elif not(continuidade):
            self.importaDadosTermica();
        
        return;
    
    def importaDadosTermica(self):
        
        # declara vetor de cvu variavel para as termicas
        self.cvu = [];
        
        val_cel = "";
        
        # index externo se refere ao indice da UHE na tabela e o interno se refere ao indice na lista do programa
        self.indexUsinaExterno = self.fonte_dados.pegaEscalar("A3", lin_offset=self.linhaOffset);
        self.sis_index = self.fonte_dados.pegaEscalar("C3", lin_offset=self.linhaOffset);
        potdisp = self.fonte_dados.pegaEscalar("F3", lin_offset=self.linhaOffset);
        self.dataMinima = self.fonte_dados.pegaEscalar("K3", lin_offset=self.linhaOffset)-1; # a termica tradicional esta no formato de 1 12
        val_cel = str(self.fonte_dados.pegaEscalar("T3", lin_offset=self.linhaOffset));
        self.inflexExistente = [];
        if (val_cel[:1] == "["):
            self.inflexExistente = val_cel[1:(len(val_cel)-1)].split(";");
            self.inflexExistente = [float(v) for v in self.inflexExistente];
            self.inflexSazonal = True;
        else:
            # o vetor eh sempre igual
            self.inflexExistente = [float(val_cel) for i in range(12)];
            self.inflexSazonal = False;

        self.cvu = self.fonte_dados.pegaVetor("AQ3", direcao='horizontal', tamanho=self.numMeses, lin_offset=self.linhaOffset);
        # repete o ultimo ano para o periodo pós
        for iper in range(self.numMesesPos):
            self.cvu.append(self.cvu[self.numMeses - 12 + iper%12]);
        self.dataSaida = self.fonte_dados.pegaEscalar("H3", lin_offset=self.linhaOffset);
        # ajusta caso nao esteja preenchido o valor data de saida
        if (self.dataSaida is None) or (isinstance(self.dataSaida, str)):
            self.dataSaida = 999;
        else:
            self.dataSaida = int(self.dataSaida)-1;
        
        if self.dataSaida > self.numMeses:
            self.dataSaida = self.numMeses;
        
        # prepara o vetor de potencias de acordo com a entrada da usina
        self.potUsina = [0 for iper in range(self.numMeses)];
        
        # ajusta a diferenca do 0 e 1 do periodo inicial
        difPer = int(self.dataMinima);
        if difPer<0:
            difPer=0;
            
        # preenche a potencia apos entrada em operacao
        for iper in range(difPer, self.dataSaida):
            self.potUsina[iper] = potdisp;
        # repete o ultimo ano para o periodo pós
        for iper in range(self.numMesesPos):
            self.potUsina.append(self.potUsina[self.numMeses-12 + iper%12]);
            
        return;
