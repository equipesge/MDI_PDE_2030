from Usina import Usina;
from RecebeDados import RecebeDados;

class UHE(Usina):

    def __init__(self, recebe_dados, offset, iUHE, nHidros, nMeses, nMesesPos, variaHidro, combSeriesHidroEol, numEol):
        
        # define fonte_dados como o objeto da classe RecebeDados e o nome da aba com as usinas UHE
        self.nomeAba = 'UHE';
        self.fonte_dados = recebe_dados;
        self.indexUsinaInterno = iUHE;
        self.numHidros = int(nHidros);
        self.numMeses = int(nMeses);
        self.numMesesPos = int(nMesesPos);
        self.offsetHidro = variaHidro;
        self.tipoCombHidroEol = combSeriesHidroEol;
        self.nSeriesEol = int(numEol);
        
        # a variavel offset e importante pq a quantidade de linhas que devem ser puladas na planilha
        self.linhaOffset = offset;
        
        # metodo referente a classe pai
        super(UHE, self).__init__(self.fonte_dados,self.nomeAba, self.linhaOffset);
        
        # importa dados especificos de UHEs
        self.importaDadosUHE();
        
        # importa especificamente as series de geracao de cada UHE
        self.serieHidrologica = [];
        self.importaSeriesHidrologicas();
        
        # importa especificamente os dados de potencia maxima disponivel
        self.potDisp = [];
        self.importaPotenciasDisp();
        
        return;
    
    def importaDadosUHE(self):
        # index externo se refere ao indice da UHE na tabela e o interno se refere ao indice na lista do programa
        self.indexUsinaExterno = self.fonte_dados.pegaEscalar("A3", lin_offset=self.linhaOffset);
        self.sis_index = self.fonte_dados.pegaEscalar("C3",  lin_offset=self.linhaOffset);
        self.potUsina = self.fonte_dados.pegaEscalar("I3",  lin_offset=self.linhaOffset); 
        self.dataMinima = self.fonte_dados.pegaEscalar("G3",  lin_offset=self.linhaOffset); # a UHE esta no formato 0 a 11
        self.ghMin = self.fonte_dados.pegaEscalar("AE3",  lin_offset=self.linhaOffset); 
        
        # pega o inicio da motorizacao e verifica se esta em motorizacao ou nao
        if (self.fonte_dados.pegaEscalar("D3",  lin_offset=self.linhaOffset) > 0):
            self.inicioMotorizacao = self.fonte_dados.pegaEscalar("G3",  lin_offset=self.linhaOffset);
        else:
            self.inicioMotorizacao =  self.fonte_dados.pegaEscalar("N3",  lin_offset=self.linhaOffset);
            
        self.nMesesMotorizacao = self.fonte_dados.pegaEscalar("K3",  lin_offset=self.linhaOffset);
        if self.nMesesMotorizacao is not None:
            # decrementa por qestao de referencia e depois pega o fim
            self.inicioMotorizacao = int(self.inicioMotorizacao) - 1;
            self.nMesesMotorizacao = int(self.nMesesMotorizacao);
        else:
            self.inicioMotorizacao = 0;
            self.nMesesMotorizacao = 0;
        return;
    
    def importaSeriesHidrologicas(self):
        # importa os dados hidrologicos separando as series de vazoes pelo offset hidrologico entrado pelo usuario na aba Geral coluna M
        
        # preenche as series hidrologicas
        for iserie in range(0, self.numHidros):
            
            # pega o vetor original de energia
            vet = self.fonte_dados.pegaVetor("AG3", direcao='horizontal', tamanho=self.numMeses, lin_offset=self.linhaOffset, col_offset=self.offsetHidro[iserie]);
            
            # percorre os periodos para fazer o ajuste da motorizacao
            for iper in range(self.numMeses):
                # se nao iniciou motorizacao zera
                if iper < self.inicioMotorizacao:
                    vet[iper] = 0;
                
                # so aplica o fator de reducao se houver motorizacao
                if (self.nMesesMotorizacao > 0):
                    f = (iper - (self.inicioMotorizacao)) / self.nMesesMotorizacao;
                    # so aplicar um fator se for redutor
                    if f < 1:
                        vet[iper] = vet[iper] * f;

            # repete o ultimo ano para o periodo pós
            for iper in range(self.numMesesPos):
                vet.append(vet[self.numMeses-12 + iper%12]);

            # insere a lista de energia ajustada na lista de hidrologias
            self.serieHidrologica.extend([vet]);

        # verifica a condicao
        if (self.tipoCombHidroEol == "completa"):
            self.serieHidrologica = self.serieHidrologica * self.nSeriesEol;

        return;
    
    def importaPotenciasDisp(self):
        
        # muda o ponteiro para a aba com as potencias disponiveis
        self.fonte_dados.defineAba('PDispUHE');
        
        # preenche as series de potencias disponiveis
        for iserie in range(0, self.numHidros):
            # pega o vetor original da aba
            vet = self.fonte_dados.pegaVetor("AG3", direcao='horizontal', tamanho=self.numMeses, lin_offset=self.linhaOffset, col_offset=self.offsetHidro[iserie]);
            
            # percorre os periodos para fazer o ajuste da motorizacao
            for iper in range(self.numMeses):
                # se nao iniciou motorizacao zera
                if iper < self.inicioMotorizacao:
                    vet[iper] = 0;
                
                # so aplica o fator de reducao se houver motorizacao
                if (self.nMesesMotorizacao > 0):
                    f = (iper - (self.inicioMotorizacao)) / self.nMesesMotorizacao;
                    # so aplicar um fator se for redutor
                    if f < 1:
                        vet[iper] = vet[iper] * f;

            # repete o ultimo ano para o periodo pós
            for iper in range(self.numMesesPos):
                vet.append(vet[self.numMeses-12 + iper%12]);
            
            # insere a lista de potencias da UHE em questao na lista geral de potencias disponiveis
            self.potDisp.extend([vet]);

        # verifica a condicao
        if (self.tipoCombHidroEol == "completa"):
            self.potDisp = self.potDisp * self.nSeriesEol;
        
        # volta com o ponteiro do XLRD para a aba UHE para nao atrapalhar a criacao dos projetos UHE
        self.fonte_dados.defineAba('UHE');
        
        return;
    
    
