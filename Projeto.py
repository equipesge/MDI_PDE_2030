from RecebeDados import RecebeDados;
from UHE import UHE;
from Termica import Termica;
from Renovavel import Renovavel;
from Usina import Usina;

class Projeto(object):
    
    def __init__(self, recebe_dados, nomeAba, offset, iProj):
        # define fonte_dados como o objeto da classe RecebeDados e recebe a aba com o tipo de projeto em questao
        self.fonte_dados = recebe_dados;
        self.fonte_dados.defineAba(nomeAba);
        self.linhaOffset = offset;
        self.indexProjeto = iProj;
        
        return;
    
class ProjetoUHE (UHE, Projeto):
    
    def __init__(self, recebe_dados, offset, iUHE, nHidros, nMeses, nMesesPos, variaHidro, combSeriesHidroEol, numEol):
        # define instancia como o objeto da classe RecebeDados e o index com os projetos de UHEs
        self.fonte_dados = recebe_dados;
        self.linhaOffset = offset;
        self.indexUsinaInterno = iUHE;
        self.numHidros = nHidros;
        self.numMeses = nMeses;
        self.numMesesPos = nMesesPos;
        self.offsetHidro = variaHidro;
        self.tipoCombHidroEol = combSeriesHidroEol;
        self.nSeriesEol = int(numEol);
        
        # metodo referente a classe pai
        super(ProjetoUHE, self).__init__(self.fonte_dados, self.linhaOffset, self.indexUsinaInterno, self.numHidros, self.numMeses, self.numMesesPos, self.offsetHidro, self.tipoCombHidroEol, self.nSeriesEol);
        
        # importa dados especificos de UHEs
        self.importaDadosProjUHE();
        
        return;
    
    def importaDadosProjUHE(self):
        # index externo se refere ao indice da UHE na tabela e o interno se refere ao indice na lista do programa
        self.indexUsinaExterno = self.fonte_dados.pegaEscalar("A3", lin_offset=self.linhaOffset);
        self.sis_index = self.fonte_dados.pegaEscalar("C3", lin_offset=self.linhaOffset);
        self.custoFixo = self.fonte_dados.pegaEscalar("E3", lin_offset=self.linhaOffset);
        self.dataMinima = self.fonte_dados.pegaEscalar("G3", lin_offset=self.linhaOffset)-1;
        return;
    
class ProjetoTermica (Termica, Projeto):
    
    def __init__(self, recebe_dados, offset, iTerm, abaTerm, nMeses, nMesesPos, continuidade):
        # define instancia como o objeto da classe RecebeDados e o nome da aba com os projetos de usinas Termicas
        self.nomeAba = abaTerm;
        self.fonte_dados = recebe_dados;
        self.linhaOffset = offset;
        self.indexUsinaInterno = iTerm;
        self.numMeses = int(nMeses);
        self.numMesesPos = int(nMesesPos);
        self.isContinua = continuidade;
        self.isProjeto = True;
        
        # metodo referente a classe pai
        super(ProjetoTermica, self).__init__(self.fonte_dados,self.nomeAba, self.linhaOffset, self.indexUsinaInterno, self.numMeses, self.numMesesPos, continuidade);
        
        # importa dados especificos de UTEs
        self.importaDadosProjTerm();
        
        return;
    
    def importaDadosProjTerm(self):
        # index externo se refere ao indice da UHE na tabela e o interno se refere ao indice na lista do programa
        self.indexUsinaExterno = self.fonte_dados.pegaEscalar("A3", lin_offset=self.linhaOffset);
        self.sis_index = self.fonte_dados.pegaEscalar("C3", lin_offset=self.linhaOffset);
        self.dataMinima = self.fonte_dados.pegaEscalar("K3", lin_offset=self.linhaOffset)-1; # o projeto de termica esta no formato 1 a 12
        self.custoFixo = self.fonte_dados.pegaEscalar("E3", lin_offset=self.linhaOffset);
        
        # verifica se eh um projeto de termica continua ou nao para definir a potencia e o cvu
        val_cel = "";
        if (self.isContinua):
            self.potUsina = 1;
            self.dataSaida = self.numMeses + self.numMesesPos;
            self.cvu = self.fonte_dados.pegaVetor("X3", direcao='horizontal', tamanho=self.numMeses, lin_offset=self.linhaOffset);
            # repete o ultimo ano para o periodo pós
            for iper in range(self.numMesesPos):
                self.cvu.append(self.cvu[self.numMeses - 12 + iper%12]);
            # para fazer a leitura da inflexiblidade deve verificar se nao eh sazonal                    
            val_cel = str(self.fonte_dados.pegaEscalar("F3", lin_offset=self.linhaOffset));
            
            # pega o fdisp
            self.fdisp = float(self.fonte_dados.pegaEscalar("P3", lin_offset=self.linhaOffset));
        else:
            self.potUsina = self.fonte_dados.pegaEscalar("F3", lin_offset=self.linhaOffset);
            self.dataSaida = self.fonte_dados.pegaEscalar("H3", lin_offset=self.linhaOffset) - 1;
            self.cvu = self.fonte_dados.pegaVetor("AQ3", direcao='horizontal', tamanho=self.numMeses, lin_offset=self.linhaOffset);
            # repete o ultimo ano para o periodo pós
            for iper in range(self.numMesesPos):
                self.cvu.append(self.cvu[self.numMeses - 12 + iper%12]);
            val_cel = str(self.fonte_dados.pegaEscalar("T3", lin_offset=self.linhaOffset));
            self.fdisp = (1-float(self.fonte_dados.pegaEscalar("L3", lin_offset=self.linhaOffset))/100)*(1-float(self.fonte_dados.pegaEscalar("M3", lin_offset=self.linhaOffset))/100);

        # se tiver colchetes eh sazonal
        self.inflexContinua = [];
        if (val_cel[:1] == "["):
            self.inflexSazonal = True;
            self.inflexContinua = val_cel[1:(len(val_cel)-1)].split(";");
            self.inflexContinua = [float(v) for v in self.inflexContinua];
        else:
            # o vetor eh sempre igual
            self.inflexContinua = [float(val_cel) for i in range(12)];
            self.inflexSazonal = False;

        return;
    
class ProjetoRenovavel (Renovavel, Projeto):
    
    def __init__(self, recebe_dados, offset, iRenov, abaRenov, numEol, combSeriesHidroEol, nHidros):
        # define instancia como o objeto da classe RecebeDados e o index com os projetos de Renovavel
        self.nomeAba = abaRenov;
        self.fonte_dados = recebe_dados;
        self.linhaOffset = offset;
        self.indexUsinaInterno = iRenov;
        self.nSeriesEol = int(numEol);
        self.tipoCombHidroEol = combSeriesHidroEol;
        self.numHidros = int(nHidros);
        
        # metodo referente a classe pai
        super(ProjetoRenovavel, self).__init__(self.fonte_dados, self.nomeAba, self.linhaOffset, self.indexUsinaInterno);
        
        # importa dados especificos de Renovaveis
        self.importaDadosProjRenovavel();

        # importa series eolicas
        self.importaSeriesEolicas();
    
        return;
    
    def importaDadosProjRenovavel(self):
        
        # index externo se refere ao indice do projeto na tabela e o interno se refere ao indice na lista do programa
        self.indexUsinaExterno = self.fonte_dados.pegaEscalar("A3", lin_offset=self.linhaOffset);
        self.sis_index = self.fonte_dados.pegaEscalar("C3", lin_offset=self.linhaOffset);
        self.custoMensal = self.fonte_dados.pegaEscalar("D3", lin_offset=self.linhaOffset);
        self.dataMinima = self.fonte_dados.pegaEscalar("E3", lin_offset=self.linhaOffset)-1;
        self.tipo = self.fonte_dados.pegaEscalar("G3", lin_offset=self.linhaOffset);

        return;

    def importaSeriesEolicas(self):

        # delaracao de variaveis usadas para importacao das series de incerteza dos ventos
        self.fatorCapacidade = [];
        self.seriesEolicas = [];
        offsetLocal = 0;
        
        # importacao de fator de capacidade e incerteza das eolicas
        if (self.tipo == "EOL"):
            self.fonte_dados.defineAba("Series Eolicas");
            # preenche as series eolicas
            while (self.fonte_dados.pegaEscalar("A2", lin_offset=offsetLocal) is not None):
                if (self.fonte_dados.pegaEscalar("A2", lin_offset=offsetLocal) == self.indexUsinaExterno):
                    for iserie in range(0, self.nSeriesEol):
                        # le valores da serie
                        vet = self.fonte_dados.pegaVetor("C2", direcao='horizontal', tamanho=12, lin_offset=offsetLocal);
                        self.seriesEolicas.extend([vet]);
                        offsetLocal += 1;
                else:
                    offsetLocal += 1;
            
            # verifica a condicao
            if (self.tipoCombHidroEol == "completa"):
                self.seriesEolicas = self.seriesEolicas*self.numHidros;
        else:
            self.fatorCapacidade = self.fonte_dados.pegaVetor("H3", "horizontal", 12, self.linhaOffset);
        
        # para garantir que o metodo termine com o ponteiro para a aba correta
        self.fonte_dados.defineAba(self.nomeAba);
        
        return;

class ProjetoReversivel (Usina, Projeto):
    
    def __init__(self, recebe_dados, offset, iReversivel, abaReversivel):
        # define instancia como o objeto da classe RecebeDados e o index com os projetos de Reversivel
        self.nomeAba = abaReversivel;
        self.fonte_dados = recebe_dados;
        self.linhaOffset = offset;
        self.indexUsinaInterno = iReversivel;
        
        # metodo referente a classe pai
        super(ProjetoReversivel, self).__init__(self.fonte_dados, self.nomeAba, self.linhaOffset);
        
        # importa dados especificos de Reversiveis
        self.importaDadosProjReversivel();
    
        return;
    
    def importaDadosProjReversivel(self):
        
        # index externo se refere ao indice do projeto na tabela e o interno se refere ao indice na lista do programa
        self.indexUsinaExterno = self.fonte_dados.pegaEscalar("A3", lin_offset=self.linhaOffset);
        self.sis_index = self.fonte_dados.pegaEscalar("C3", lin_offset=self.linhaOffset);
        self.custoMensal = self.fonte_dados.pegaEscalar("D3", lin_offset=self.linhaOffset);
        self.rendimento = self.fonte_dados.pegaEscalar("F3", lin_offset=self.linhaOffset);
        self.dataMinima = self.fonte_dados.pegaEscalar("E3", lin_offset=self.linhaOffset)-1;
        return;