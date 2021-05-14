from RecebeDados import RecebeDados;
from UHE import UHE;
from Termica import Termica;
from Agrint import Agrint;
from Projeto import *;
from datetime import *;
from Restricoes import *;

class Sistema:
    
    def __init__(self, recebe_dados, combSeriesHidroEol):
        
        # define fonte_dados como o objeto da classe RecebeDados
        self.fonte_dados = recebe_dados;

        # recebe a opcao marcada no checkbox do excel para saber o numero de condicoes a criar
        self.tipoCombHidroEol = combSeriesHidroEol;
        
        # importa os dados gerais do sistema
        self.importaDadosGerais();
        
        # cria as condicoes
        self.criaCondicoes();
        
        # preenche os subsistemas
        self.subsistemas = [Subsistema(self.fonte_dados, isis+1, self.numAnos, self.numMesesTotal, self.nsis, self.numCondicoes, self.nPatamares) for isis in range(0, self.nsis)];

        # cria os diferentes tipos de usinas e importa as informacoes    
        self.importaUsinas();
        
        # totaliza as series de energia existentes
        self.totalizaSeries();

        # importa a duracao e a carga dos patamares
        self.importaDurPatamar();
        
        # importa os fatores de Carga para as usinas
        self.importafatorPat();

        # importa o indice dos patamares em que nao pode haver geracao e bombeamento, respectivamente
        self.importaRestReversiveis();
        
        # cria os agrupamentos de interligacoes
        self.criaAgrints();
        
        # monta listas gerais - independentes de subsistemas
        self.montaListasGerais();

        # faz a leitura das restricoes adicionais
        self.addRestricoes();
        
        return;
    
    def importaDadosGerais(self):

        # muda para a aba dos patamares para importar o numero de patamares porque sera usado na criacao dos subsistemas
        self.fonte_dados.defineAba(nomeAba = 'Patamar');

        # importa o numero de patamares
        self.nPatamares = int(self.fonte_dados.pegaEscalar("D2"));

        # o custo de deficit passa a variar com o patamar
        self.custoDefc = [];
        self.custoDefc.extend(self.fonte_dados.pegaVetor("I2", direcao='horizontal', tamanho=self.nPatamares));
        
        # define a aba da planilha ser usada para importar os outros dados gerais
        self.fonte_dados.defineAba(nomeAba = 'GERAL');
        
        # carrega parametros gerais do sistema
        self.taxaDesc = float(self.fonte_dados.pegaEscalar("D5")); # taxa de desconto
        self.inicioSim = self.fonte_dados.pegaEscalar("G3");  # inicio da simulacao
        self.numMeses = int(self.fonte_dados.pegaEscalar("G5")); # numero de meses na simulacao
        self.numAnos = int(self.numMeses/12); # numero de anos na simulacao
        self.numHidros = int(self.fonte_dados.pegaEscalar("G6")); # numero de series hidrologicas consideradas
        self.horasMes = self.fonte_dados.pegaEscalar("G7"); # numero de horas no mes
        self.perMinExpT = int(self.fonte_dados.pegaEscalar("G8")); # periodo minimo para a expansao da transmissao
        self.numEol = int(self.fonte_dados.pegaEscalar("G9")); # numero de series eolicas consideradas
        self.nsis = int(self.fonte_dados.pegaEscalar("G10")); # numero de subsistemas
        self.custoDefPot = self.fonte_dados.pegaEscalar("C8"); # custo de deficit de potencia
        self.restPot = self.fonte_dados.pegaEscalar("C9"); # demanda adicional para a restrição de potência
        self.pldMin = self.fonte_dados.pegaEscalar("C10"); # valor da energia para reversiveis
        self.numMesesPos = int(self.fonte_dados.pegaEscalar("G12"))*12; # numero de meses pos na simulacao

        # define numero de meses total, somando o pós
        self.numMesesTotal = self.numMeses + self.numMesesPos;
        
        # sao retirados 2 dias para compatibilizar os formatos de data da classe do python com o excel - mudanca manual 
        dataBase = date(1900, 1, 1);
        td = timedelta(days = (int(self.inicioSim)-2));
        self.inicioSim = dataBase + td;
        
        # computa o ano inicial
        self.anoInicial = self.inicioSim.year;
        
        # pega a variacao de colunas necessaria para cada serie hidrologica
        self.offsetHidro = [];
        self.offsetHidro.extend(self.fonte_dados.pegaVetor("M2", direcao='vertical', tamanho=self.numHidros));
        self.probHidro = self.fonte_dados.pegaVetor("K2", direcao='vertical', tamanho=self.numHidros);
        
        return;

    def importaUsinas(self):
        # cria as variaveis contadoras indicando os index de cada usina
        iUsina = 0;
        iUHE = 0;
        iprojUHE = 0;
        iTerm = 0;
        iProjTerm = 0;
        iprojRenov = 0;
        iprojReversivel = 0;
        
        # define a aba UHE
        self.fonte_dados.defineAba("UHE");
        
        # percorre a aba e aloca as usinas UHE em uma lista
        while (self.fonte_dados.pegaEscalar("A3", lin_offset=iUsina)is not None):
            
            # verifica se eh UHE (== 0) ou projeto de UHE (!=0)
            if (self.fonte_dados.pegaEscalar("D3",lin_offset=iUsina) == 0):
                
                # instancia a usina e aloca no respectivo subsistema
                usinaUHE = UHE(self.fonte_dados, iUsina, iUHE, self.numHidros, self.numMeses, self.numMesesPos,  self.offsetHidro, self.tipoCombHidroEol, self.numEol);
            
                # -1 eh referente ao fato de subsistemas comecarem do 0
                self.subsistemas[int(usinaUHE.sis_index)-1].addUsinaUHE(usinaUHE);
                
                # incrementa o contador
                iUHE+=1;
                
            elif (self.fonte_dados.pegaEscalar("D3",lin_offset=iUsina) > 0):
                # instancia o Projeto de UHE e aloca no respectivo subsistema
                projUHE = ProjetoUHE(self.fonte_dados, iUsina, iprojUHE, self.numHidros, self.numMeses, self.numMesesPos, self.offsetHidro, self.tipoCombHidroEol, self.numEol);

                # -1 eh referente ao fato de subsistemas comecarem do 0
                self.subsistemas[int(projUHE.sis_index)-1].addProjetoUHE(projUHE);
                
                # incrementa o contador de projetos de usinas termicas
                iprojUHE+=1;
                
            # incrementa o contador de usinas na aba
            iUsina+=1;    
    
        # define a aba TERM
        self.fonte_dados.defineAba("TERM");
        iUsina = 0;
        
        # percorre a aba TERM e aloca as usinas Termicas em uma lista
        while (self.fonte_dados.pegaEscalar("D3", lin_offset=iUsina)is not None):
            
            # verifica se eh termica (== 0) ou projeto de termica (!=0)
            if (self.fonte_dados.pegaEscalar("D3",lin_offset=iUsina) == 0):
                # instancia a usina e aloca no respectivo subsistema
                usinaTermica = Termica(self.fonte_dados, "TERM", iUsina, iTerm, self.numMeses, self.numMesesPos);

                # -1 eh referente ao fato de subsistemas comecarem do 0
                self.subsistemas[int(usinaTermica.sis_index)-1].addUsinaTermica(usinaTermica);
                
                # incrementa o contador de usinas termicas
                iTerm+=1;
            
            elif (self.fonte_dados.pegaEscalar("D3",lin_offset=iUsina) > 0):
                # instancia o Projeto de Termica e aloca no respectivo subsistema
                projTermica = ProjetoTermica(self.fonte_dados, iUsina, iProjTerm, "TERM", self.numMeses, self.numMesesPos, False);

                # -1 eh referente ao fato de subsistemas comecarem do 0
                self.subsistemas[int(projTermica.sis_index)-1].addProjetoTermica(projTermica);
                
                # incrementa o contador de projetos de usinas termicas
                iProjTerm+=1;
                
            # incrementa o contador de usinas na aba
            iUsina+=1;
    
        self.fonte_dados.defineAba("TermicasContinuas");
        iUsina = 0;
        
        # percorre a aba TermicasContinuas e aloca as usinas Termicas em uma lista
        while (self.fonte_dados.pegaEscalar("A3", lin_offset=iUsina)is not None):
            
            # incrementa o contador de projetos primeiro pq ja existe um projeto de termica anterior
            iProjTerm+=1;
            
            # instancia o Projeto de Termica e aloca no respectivo subsistema
            projTermica = ProjetoTermica(self.fonte_dados, iUsina, iProjTerm, "TermicasContinuas", self.numMeses, self.numMesesPos, True);

            # -1 eh referente ao fato de subsistemas comecarem do 0
            self.subsistemas[int(projTermica.sis_index)-1].addProjetoTermica(projTermica);
            
            # incrementa o contador de usinas na aba
            iUsina+=1;
    
        self.fonte_dados.defineAba("Renov Ind.");
        iUsina = 0;
        
        # percorre a aba EOL Projetos e aloca os projetos de Renovavel em uma lista
        while (self.fonte_dados.pegaEscalar("A3", lin_offset=iUsina)is not None):
            
            # instancia o Projeto de Termica e aloca no respectivo subsistema
            projRenovavel = ProjetoRenovavel(self.fonte_dados, iUsina, iprojRenov, "Renov Ind.", self.numEol, self.tipoCombHidroEol, self.numHidros);

            # -1 eh referente ao fato de subsistemas comecarem do 0
            self.subsistemas[int(projRenovavel.sis_index)-1].addProjetoRenovavel(projRenovavel);
            
            # incrementa o contador de usinas na aba
            iUsina+=1;
            
            # incrementa o contador de projetos Renovaveis
            iprojRenov+=1;

        self.fonte_dados.defineAba("Armazenamento");
        iUsina = 0;
        
        # percorre a aba Reversivel e aloca os projetos de Reversivel em uma lista
        while (self.fonte_dados.pegaEscalar("A3", lin_offset=iUsina)is not None):
            
            # instancia o Projeto de Reversivel e aloca no respectivo subsistema
            projReversivel = ProjetoReversivel(self.fonte_dados, iUsina, iprojReversivel, "Armazenamento");

            # -1 eh referente ao fato de subsistemas comecarem do 0
            self.subsistemas[int(projReversivel.sis_index)-1].addProjetoReversivel(projReversivel);
            
            # incrementa o contador de usinas na aba
            iUsina+=1;
            
            # incrementa o contador de projetos Reversiveis
            iprojReversivel+=1;
            
        # loop para importar dados das usinas renovaveis existentes continuas e alocar nos subsistemas
        self.fonte_dados.defineAba("UNSI");
        iUsina = 0;
            
        # percorre a aba passando por cada Renovavel Existente
        while (self.fonte_dados.pegaEscalar("A3", lin_offset=iUsina)is not None):
        
            # verifica a data de entrada da renovavel decidida
            per_entrada = int(self.fonte_dados.pegaEscalar("G3", lin_offset=iUsina));

            # ajuste por conta da referencia inicial
            if per_entrada == 0:
                per_entrada = 1;
            
            # loop para acrescentar os dados de potencial mensalmente na lista montanteRenovavel a partir do periodo de entrada
            for iper in range(per_entrada-1,self.numMesesTotal):
                imes = iper%12;
                
                # incrementa as listas com os montantes de energia e potencia fornecidas por cada tipo de usina
                if (str(self.fonte_dados.pegaEscalar("D3", lin_offset=iUsina)) == "BIO"):
                    self.subsistemas[int(self.fonte_dados.pegaEscalar("C3", lin_offset=iUsina))-1].montanteRenovExBIO[iper] += self.fonte_dados.pegaEscalar("H3", lin_offset=iUsina, col_offset=imes);
                    self.subsistemas[int(self.fonte_dados.pegaEscalar("C3", lin_offset=iUsina))-1].montanteRenovExBIOPot[iper] += self.fonte_dados.pegaEscalar("E3", lin_offset=iUsina);
                if (str(self.fonte_dados.pegaEscalar("D3", lin_offset=iUsina)) == "SOL"):
                    self.subsistemas[int(self.fonte_dados.pegaEscalar("C3", lin_offset=iUsina))-1].montanteRenovExUFV[iper] += self.fonte_dados.pegaEscalar("H3", lin_offset=iUsina, col_offset=imes);
                    self.subsistemas[int(self.fonte_dados.pegaEscalar("C3", lin_offset=iUsina))-1].montanteRenovExUFVPot[iper] += self.fonte_dados.pegaEscalar("E3", lin_offset=iUsina);
                if (str(self.fonte_dados.pegaEscalar("D3", lin_offset=iUsina)) == "EOL"):
                    self.subsistemas[int(self.fonte_dados.pegaEscalar("C3", lin_offset=iUsina))-1].montanteRenovExEOL[iper] += self.fonte_dados.pegaEscalar("H3", lin_offset=iUsina, col_offset=imes);
                    self.subsistemas[int(self.fonte_dados.pegaEscalar("C3", lin_offset=iUsina))-1].montanteRenovExEOLPot[iper] += self.fonte_dados.pegaEscalar("E3", lin_offset=iUsina);
                if (str(self.fonte_dados.pegaEscalar("D3", lin_offset=iUsina)) == "PCH"):
                    self.subsistemas[int(self.fonte_dados.pegaEscalar("C3", lin_offset=iUsina))-1].montanteRenovExPCH[iper] += self.fonte_dados.pegaEscalar("H3", lin_offset=iUsina, col_offset=imes);
                    self.subsistemas[int(self.fonte_dados.pegaEscalar("C3", lin_offset=iUsina))-1].montanteRenovExPCHPot[iper] += self.fonte_dados.pegaEscalar("E3", lin_offset=iUsina); 
            
            # incrementa o contador de usinas na aba
            iUsina+=1;
    
    # atentar para o fato de que existem dois metodos totalizaSeries, um na classe Sistema e um na classe Subsistema
    def totalizaSeries(self):
        # percorre os subsistemas e manda cada um totalizar suas series hidrologicas
        for subsis in self.subsistemas:
            subsis.totalizaSeries();
        return;

    # atraves da opcao marcada no checkbox do excel, estabelece o numero de condicoes misturando series hidrologicas e eolicas
    def criaCondicoes(self):

        # inicializa a variavel condicoes
        self.numCondicoes = 0;

        if (self.tipoCombHidroEol == "completa"):
            self.numCondicoes = self.numHidros * self.numEol; 
        elif (self.tipoCombHidroEol == "intercalada"):
            self.numCondicoes = self.numHidros;
        else:
            print("opcao de combinacao de series hidrologicas com eolicas nao marcada");
        
        return;

    def importaDurPatamar(self):
        
        # muda a aba para a que possui os patamares
        self.fonte_dados.defineAba('Patamar');
        
        # declara o vetor de duracao dos patamares
        # importante ressaltar que a carga dos patamares varia por subsistema portanto o vetor carga foi declarado na classe subsistema
        self.duracaoPatamar = [[0 for x in range(0,self.numMeses)] for y in range(0, self.nPatamares)];

        # importa os valores da duracao de cada patamar
        linhaOffset = 0;
        for patamar in range(0, self.nPatamares):
            self.duracaoPatamar[patamar] = self.fonte_dados.pegaVetor("C6", "horizontal", self.numMeses, linhaOffset);
            linhaOffset += 1;
            # repete o ultimo ano para o periodo pós
            for iper in range(self.numMesesPos):
                self.duracaoPatamar[patamar].append(self.duracaoPatamar[patamar][self.numMeses - 12 + iper%12]);
        
        return;
    
    def importafatorPat(self):
        
        # muda para a aba correta
        self.fonte_dados.defineAba('Fatores');

        # define fatores igual a 1 para todos os subsistemas e patamares
        for sis in range(self.nsis):
            for pat in range(self.nPatamares):
                for per in range(12):
                    self.subsistemas[sis].fatorPatEOL[pat][per] = 1;
                    self.subsistemas[sis].fatorPatEOLEx[pat][per] = 1;
                    self.subsistemas[sis].fatorPatEOF[pat][per] = 1;
                    self.subsistemas[sis].fatorPatUFV[pat][per] = 1;
                    self.subsistemas[sis].fatorPatUFVEx[pat][per] = 1;
                    self.subsistemas[sis].fatorPatBIO[pat][per] = 1;
                    self.subsistemas[sis].fatorPatPCH[pat][per] = 1;
        
        # contador para auxiliar no offset das linhas e colunas
        linhaOffset = 0;
        colunaOffset = 0;
        
        # preenche os vetores com os fatores para as renovaveis eolicas indicativas
        cel = str(self.fonte_dados.pegaEscalar("A1", lin_offset=linhaOffset))
        while (cel.startswith("EOL IND") == False):
            linhaOffset += 1;
            cel = str(self.fonte_dados.pegaEscalar("A1", lin_offset=linhaOffset))
        linhaOffset += 2;
        while (self.fonte_dados.pegaEscalar("A1", lin_offset=linhaOffset) is not None):
            # identifica o subsistema e le os fatores
            subsis = self.subsistemas[int(self.fonte_dados.pegaEscalar('A1', linhaOffset))-1];

            for patamar in range(0, self.nPatamares):

                if patamar>0:
                    colunaOffset = 12*patamar;
                
                subsis.fatorPatEOL[patamar] = self.fonte_dados.pegaVetor('B1', 'horizontal', 12, linhaOffset, colunaOffset);
            # incrementa o contador de linha e zera o contador de coluna
            linhaOffset += 1;
            colunaOffset = 0;

        #zera os contadores para procurar a próxima fonte
        linhaOffset = 0;
        colunaOffset = 0;

        # preenche os vetores com os fatores para as renovaveis eolicas existentes e contratadas
        cel = str(self.fonte_dados.pegaEscalar("A1", lin_offset=linhaOffset))
        while (cel.startswith("EOL EX") == False):
            linhaOffset += 1;
            cel = str(self.fonte_dados.pegaEscalar("A1", lin_offset=linhaOffset))
        linhaOffset += 2;
        while (self.fonte_dados.pegaEscalar("A1", lin_offset=linhaOffset) is not None):
            # identifica o subsistema e le os fatores
            subsis = self.subsistemas[int(self.fonte_dados.pegaEscalar('A1', linhaOffset))-1];

            for patamar in range(0, self.nPatamares):

                if patamar>0:
                    colunaOffset = 12*patamar;
                
                subsis.fatorPatEOLEx[patamar] = self.fonte_dados.pegaVetor('B1', 'horizontal', 12, linhaOffset, colunaOffset);
            # incrementa o contador de linha e zera o contador de coluna
            linhaOffset += 1;
            colunaOffset = 0;

        #zera os contadores para procurar a próxima fonte
        linhaOffset = 0;
        colunaOffset = 0;

        # preenche os vetores com os fatores para as renovaveis eolicas offshore
        cel = str(self.fonte_dados.pegaEscalar("A1", lin_offset=linhaOffset))
        while (cel.startswith("EOL OF") == False):
            linhaOffset += 1;
            cel = str(self.fonte_dados.pegaEscalar("A1", lin_offset=linhaOffset))
        linhaOffset += 2;
        while (self.fonte_dados.pegaEscalar("A1", lin_offset=linhaOffset) is not None):
            # identifica o subsistema e le os fatores
            subsis = self.subsistemas[int(self.fonte_dados.pegaEscalar('A1', linhaOffset))-1];

            for patamar in range(0, self.nPatamares):

                if patamar>0:
                    colunaOffset = 12*patamar;
                
                subsis.fatorPatEOF[patamar] = self.fonte_dados.pegaVetor('B1', 'horizontal', 12, linhaOffset, colunaOffset);
            # incrementa o contador de linha e zera o contador de coluna
            linhaOffset += 1;
            colunaOffset = 0;

        #zera os contadores para procurar a próxima fonte
        linhaOffset = 0;
        colunaOffset = 0;

        # preenche os vetores com os fatores para as renovaveis fotovoltaicas indicativas
        cel = str(self.fonte_dados.pegaEscalar("A1", lin_offset=linhaOffset))
        while (cel.startswith("UFV IND") == False):
            linhaOffset += 1;
            cel = str(self.fonte_dados.pegaEscalar("A1", lin_offset=linhaOffset))
        linhaOffset += 2;
        while (self.fonte_dados.pegaEscalar("A1", lin_offset=linhaOffset) is not None):
            # identifica o subsistema e le os fatores
            subsis = self.subsistemas[int(self.fonte_dados.pegaEscalar('A1', linhaOffset))-1];

            for patamar in range(0, self.nPatamares):

                if patamar>0:
                    colunaOffset = 12*patamar;
                
                subsis.fatorPatUFV[patamar] = self.fonte_dados.pegaVetor('B1', 'horizontal', 12, linhaOffset, colunaOffset);
            # incrementa o contador de linha e zera o contador de coluna
            linhaOffset += 1;
            colunaOffset = 0;

        #zera os contadores para procurar a próxima fonte
        linhaOffset = 0;
        colunaOffset = 0;

        # preenche os vetores com os fatores para as renovaveis fotovoltaicas existentes e contratadas
        cel = str(self.fonte_dados.pegaEscalar("A1", lin_offset=linhaOffset))
        while (cel.startswith("UFV EX") == False):
            linhaOffset += 1;
            cel = str(self.fonte_dados.pegaEscalar("A1", lin_offset=linhaOffset))
        linhaOffset += 2;
        while (self.fonte_dados.pegaEscalar("A1", lin_offset=linhaOffset) is not None):
            # identifica o subsistema e le os fatores
            subsis = self.subsistemas[int(self.fonte_dados.pegaEscalar('A1', linhaOffset))-1];

            for patamar in range(0, self.nPatamares):

                if patamar>0:
                    colunaOffset = 12*patamar;
                
                subsis.fatorPatUFVEx[patamar] = self.fonte_dados.pegaVetor('B1', 'horizontal', 12, linhaOffset, colunaOffset);
            # incrementa o contador de linha e zera o contador de coluna
            linhaOffset += 1;
            colunaOffset = 0;

        #zera os contadores para procurar a próxima fonte
        linhaOffset = 0;
        colunaOffset = 0;

        # preenche os vetores com os fatores para as renovaveis biomassa
        cel = str(self.fonte_dados.pegaEscalar("A1", lin_offset=linhaOffset))
        while (cel.startswith("BIO") == False):
            linhaOffset += 1;
            cel = str(self.fonte_dados.pegaEscalar("A1", lin_offset=linhaOffset))
        linhaOffset += 2;
        while (self.fonte_dados.pegaEscalar("A1", lin_offset=linhaOffset) is not None):
            # identifica o subsistema e le os fatores
            subsis = self.subsistemas[int(self.fonte_dados.pegaEscalar('A1', linhaOffset))-1];

            for patamar in range(0, self.nPatamares):

                if patamar>0:
                    colunaOffset = 12*patamar;
                
                subsis.fatorPatBIO[patamar] = self.fonte_dados.pegaVetor('B1', 'horizontal', 12, linhaOffset, colunaOffset);
            # incrementa o contador de linha e zera o contador de coluna
            linhaOffset += 1;
            colunaOffset = 0;

        #zera os contadores para procurar a próxima fonte
        linhaOffset = 0;
        colunaOffset = 0;

        # preenche os vetores com os fatores para as renovaveis PCH
        cel = str(self.fonte_dados.pegaEscalar("A1", lin_offset=linhaOffset))
        while (cel.startswith("PCH") == False):
            linhaOffset += 1;
            cel = str(self.fonte_dados.pegaEscalar("A1", lin_offset=linhaOffset))
        linhaOffset += 2;
        while (self.fonte_dados.pegaEscalar("A1", lin_offset=linhaOffset) is not None):
            # identifica o subsistema e le os fatores
            subsis = self.subsistemas[int(self.fonte_dados.pegaEscalar('A1', linhaOffset))-1];

            for patamar in range(0, self.nPatamares):

                if patamar>0:
                    colunaOffset = 12*patamar;
                
                subsis.fatorPatPCH[patamar] = self.fonte_dados.pegaVetor('B1', 'horizontal', 12, linhaOffset, colunaOffset);
            # incrementa o contador de linha e zera o contador de coluna
            linhaOffset += 1;
            colunaOffset = 0;

        #zera os contadores para procurar a próxima fonte
        linhaOffset = 0;
        colunaOffset = 0;
          
        return;

    def importaRestReversiveis(self):

        # muda para a aba das usinas reversiveis para importar os patamares em que nao pode haver geracao ou bombeamento
        self.fonte_dados.defineAba(nomeAba = 'Armazenamento');
        
        # declara os vetores que armazenarao os patamares
        self.naoGeraReversivel = [];
        self.naoBombReversivel = []; 

        # importa os patamares com restricao de bombeamento
        iReversivel = 0;
        while (self.fonte_dados.pegaEscalar("H3", lin_offset=iReversivel)is not None):
           
            self.naoBombReversivel.append(int(self.fonte_dados.pegaEscalar("H3", lin_offset=iReversivel))-1); 
            
            # incrementa o offset
            iReversivel+=1;

        # importa os patamares com restricao de geracao
        iReversivel = 0;
        while (self.fonte_dados.pegaEscalar("I3", lin_offset=iReversivel)is not None):
           
            self.naoGeraReversivel.append(int(self.fonte_dados.pegaEscalar("I3", lin_offset=iReversivel))-1); 
            
            # incrementa o offset
            iReversivel+=1;

        return;
    
    def criaAgrints(self):
        
        # define a aba do excel em que os agrints estao localizados e declara variavel
        self.fonte_dados.defineAba("AGRINT-Grupos");
        num_agrints = 0;
        
        #descobre quantos agrints tem na tabela
        while (self.fonte_dados.pegaEscalar("A2", lin_offset = num_agrints) is not None):
            num_agrints += 1;
        
        self.agrints = [Agrint(self.fonte_dados, iagrint, self.numMeses, self.numMesesPos, self.nPatamares) for iagrint in range(0, num_agrints)]; 
        
        # monta a lista para impressao
        lista = [a.construirLista(self.nsis) for a in self.agrints];
        
        return;
        
    def montaListasGerais(self):
        # preenche hashs com todas as usinas de cada tipo independentemente de subsistemas
        self.listaGeralProjUHE = {usina.nomeUsina : usina for isis in range(0, self.nsis) for usina in self.subsistemas[isis].listaProjUHE};
        self.listaGeralProjTerm = {usina.nomeUsina : usina for isis in range(0, self.nsis) for usina in self.subsistemas[isis].listaProjTermica};
        self.listaGeralProjRenov = {usina.nomeUsina : usina for isis in range(0, self.nsis) for usina in self.subsistemas[isis].listaProjRenovavel};
        self.listaGeralProjReversivel = {usina.nomeUsina : usina for isis in range(0, self.nsis) for usina in self.subsistemas[isis].listaProjReversivel};
        self.listaGeralTerm = {usina.nomeUsina : usina for isis in range(0, self.nsis) for usina in self.subsistemas[isis].listaTermica};

        # preenche de novo por numero
        self.listaIndGeralProjUHE = {usina.indexUsinaExterno : usina for isis in range(0, self.nsis) for usina in self.subsistemas[isis].listaProjUHE};
        self.listaIndGeralProjTerm = {usina.indexUsinaExterno : usina for isis in range(0, self.nsis) for usina in self.subsistemas[isis].listaProjTermica};
        self.listaIndGeralProjRenov = {usina.indexUsinaInterno : usina for isis in range(0, self.nsis) for usina in self.subsistemas[isis].listaProjRenovavel};
        self.listaIndGeralProjReversivel = {usina.indexUsinaInterno : usina for isis in range(0, self.nsis) for usina in self.subsistemas[isis].listaProjReversivel};

        return;

    def addRestricoes(self):
        # chama o controlador de restricoes
        self.restricoes = Restricoes(self.fonte_dados, self.numAnos, self);
    
class Subsistema:
    
        def __init__(self, recebe_dados, sis_index, nanos, nmesesTotal, nsis, nCondicoes, numPatamares):
            # define o objeto da classe RecebeDados
            self.fonte_dados = recebe_dados;
            self.sis_index = sis_index;
            self.numAnos = nanos;
            self.numMeses = nanos*12;
            self.numMesesTotal = nmesesTotal;
            self.numMesesPos = nmesesTotal - nanos*12;
            self.nsis = nsis;
            self.numCondicoes = nCondicoes;
            self.nPatamares = numPatamares;
            
            # declaracao das listas de usinas e de projetos de usinas
            self.listaUHE = [];
            self.listaTermica = [];
            self.listaProjUHE = [];
            self.listaProjTermica = [];
            self.listaProjReversivel = [];
            self.listaProjRenovavel = []; # lista geral com todos os projetos renovaveis
            self.listaProjBIO = []; # lista para renovavel biomassa
            self.listaProjUFV = []; # lista para renovaveis solares
            self.listaProjEOL = []; # lista para renovaveis eolicas
            self.listaProjEOF = []; # lista para renovaveis eolicas offshore
            self.listaProjPCH = []; # lista para renovaveis PCH

            # declara o vetor bidimensional com a profundidade de cada patamar por subsistema
            self.cargaPatamar = [[0 for iper in range(0,self.numMeses)] for ipat in range(0, self.nPatamares)];
            
            # declara os vetores com os fatores de ponta para cada tipo de usina
            self.fatorPatUFV = [[0 for iper in range(0,12)] for ipat in range(0, self.nPatamares)];
            self.fatorPatEOL = [[0 for iper in range(0,12)] for ipat in range(0, self.nPatamares)];
            self.fatorPatEOF = [[0 for iper in range(0,12)] for ipat in range(0, self.nPatamares)];
            self.fatorPatBIO = [[0 for iper in range(0,12)] for ipat in range(0, self.nPatamares)];
            self.fatorPatPCH = [[0 for iper in range(0,12)] for ipat in range(0, self.nPatamares)];
            self.fatorPatEOLEx = [[0 for iper in range(0,12)] for ipat in range(0, self.nPatamares)];
            self.fatorPatUFVEx = [[0 for iper in range(0,12)] for ipat in range(0, self.nPatamares)];
            
            # declara e inicializa a lista com o montante de energia de UNSI existentes por mes
            self.montanteRenovExBIO = [0 for iper in range(0,self.numMesesTotal)];
            self.montanteRenovExEOL = [0 for iper in range(0,self.numMesesTotal)];
            self.montanteRenovExUFV = [0 for iper in range(0,self.numMesesTotal)];
            self.montanteRenovExPCH = [0 for iper in range(0,self.numMesesTotal)];

             # declara e inicializa a lista com o montante de potencia de UNSI existentes por mes
            self.montanteRenovExBIOPot = [0 for iper in range(0,self.numMesesTotal)];
            self.montanteRenovExEOLPot = [0 for iper in range(0,self.numMesesTotal)];
            self.montanteRenovExUFVPot = [0 for iper in range(0,self.numMesesTotal)];
            self.montanteRenovExPCHPot = [0 for iper in range(0,self.numMesesTotal)];

            # metodo para incorporar a carga de cada patamar
            self.importaCargaPatamar();
            
            # metodo para incorporar os dados da demanda ao modelo
            self.importaDemanda();
            
            # metodo para importar dados de interligacao
            self.importaInterligacoes();
            
            return;

        def importaCargaPatamar(self):
        
            # muda a aba para a que possui os patamares
            self.fonte_dados.defineAba('Patamar');

            # obriga ele a pular as linhas ate os dados do proximo subsistema
            linhaOffset = 32 + (4*(self.sis_index-1)) + (((59 - 33) - self.nPatamares)*(self.sis_index-1));

            # importa os valores da carga de cada patamar para cada subsistema
            for ipat in range (0, self.nPatamares):
                self.cargaPatamar[ipat] = self.fonte_dados.pegaVetor("C1", "horizontal", self.numMeses, linhaOffset);
                linhaOffset += 1;
        
            return;
                
        def importaDemanda(self):
            # declaracao das listas de demanda por subsistema
            self.demandaEnerg = [[0 for iper in range(0,self.numMeses)] for ipat in range(0, self.nPatamares)]; 
            self.demandaMedia = []; # vetor auxiliar para fazer a operacao
            
            # seta a aba demanda de Energia
            self.fonte_dados.defineAba(nomeAba = 'Demanda NW');
            
            # percorre os anos e preenche o vetor auxiliar com a demanda
            for iano in range(0, int(self.numAnos)):
                self.demandaMedia.extend(self.fonte_dados.pegaVetor("C4", direcao='horizontal', tamanho=12, lin_offset=iano+(self.sis_index-1)*(self.numAnos+2)));
            
            # percorre os patamares e preenche demandaEnerg multiplicando a demanda pelos fatores de cada patamar em cada periodo
            for ipat in range (0, int(self.nPatamares)):
                self.demandaEnerg[ipat] = [self.demandaMedia[i] * self.cargaPatamar[ipat][i] for i in range(len(self.demandaMedia))];
                # repete o ultimo ano para o periodo pós
                for iper in range(self.numMesesPos):
                    self.demandaEnerg[ipat].append(self.demandaEnerg[ipat][self.numMeses - 12 + iper%12]);

            return;

        def importaInterligacoes(self):
            # inicializa os vetores
            self.capExistente = [[[0 for iper in range(0,self.numMesesTotal)] for ipat in range(0,self.nPatamares)] for isis in range(0,self.nsis)];
            self.custoExpansao = [0 for isis in range(0,self.nsis)];
            self.limiteInterc = [1 for isis in range(0,self.nsis)];
            self.perdasInterc = [[[0 for iper in range(0,self.numMesesTotal)] for ipat in range(0,self.nPatamares)] for isis in range(0,self.nsis)];
            
            linhaOffset = 0;
            colunaOffset = 0;
            
            # define a aba da planilha em que estao os custos de expansao
            self.fonte_dados.defineAba("Exp Interc");
            
            # carrega os valores de custo de expansao
            for jsis in range(0, self.nsis):
                self.custoExpansao[jsis] = self.fonte_dados.pegaEscalar("B2", lin_offset=self.sis_index, col_offset = colunaOffset);
                colunaOffset+=1;

            colunaOffset = 0;

            # define a aba da planilha em que estao as perdas nas interligações
            self.fonte_dados.defineAba("Perdas");
            
            # carrega os valores de perdas
            while (self.fonte_dados.pegaEscalar("B2", lin_offset=linhaOffset) != self.sis_index):
                linhaOffset += 1;
            while (self.fonte_dados.pegaEscalar("B2", lin_offset=linhaOffset) == self.sis_index):
                for ipat in range(0, self.nPatamares):
                    sis_para = int(self.fonte_dados.pegaEscalar("C2", lin_offset=linhaOffset))
                    self.perdasInterc[sis_para - 1][ipat] = self.fonte_dados.pegaVetor("D2", "horizontal", self.numMeses, linhaOffset);
                    # repete o ultimo ano para o periodo pós
                    for iper in range(self.numMesesPos):
                        self.perdasInterc[sis_para - 1][ipat].append(self.perdasInterc[sis_para - 1][ipat][self.numMeses - 12 + iper%12]);
                    linhaOffset += 1;

            # define a aba da planilha em que estao os intercambios
            self.fonte_dados.defineAba("Intercambios");

            linhaOffset = 0;

            # carrega os valores de capacidade existente - self.sis_index + 1 = indexInterno + 1 = indexExterno
            while (self.fonte_dados.pegaEscalar("B2", lin_offset=linhaOffset) != self.sis_index):
                linhaOffset += 1;
            while (self.fonte_dados.pegaEscalar("B2", lin_offset=linhaOffset) == self.sis_index):
                for ipat in range(0, self.nPatamares):
                    sis_para = int(self.fonte_dados.pegaEscalar("C2", lin_offset=linhaOffset))
                    self.capExistente[sis_para - 1][ipat] = self.fonte_dados.pegaVetor("D2", "horizontal", self.numMeses, linhaOffset);
                    # repete o ultimo ano para o periodo pós
                    for iper in range(self.numMesesPos):
                        self.capExistente[sis_para - 1][ipat].append(self.capExistente[sis_para - 1][ipat][self.numMeses - 12 + iper%12]);
                    linhaOffset += 1;

            # define a aba da planilha em que estao os limites de intercambio
            self.fonte_dados.defineAba("LimiteInterc");

            # carrega os valores de limitaçao de intercambio para cada subsistema
            colunaOffset = 0;
            for jsis in range(0, self.nsis):
                self.limiteInterc[jsis] = self.fonte_dados.pegaEscalar("B1", lin_offset=self.sis_index, col_offset = colunaOffset);
                colunaOffset+=1;
            
            # finaliza o metodo
            return;
        
        def addUsinaUHE(self, usina):
            self.listaUHE.append(usina);
            return;
        
        def addUsinaTermica(self, usina):
            self.listaTermica.append(usina);
            return;
        
        def addProjetoUHE(self, projeto):
            self.listaProjUHE.append(projeto);
            return;
        
        def addProjetoTermica(self, projeto):
            self.listaProjTermica.append(projeto);
            return;

        def addProjetoReversivel(self, projeto):
            self.listaProjReversivel.append(projeto);
            return;
        
        def addProjetoRenovavel(self, projeto):
            self.listaProjRenovavel.append(projeto);

            if (projeto.tipo == "BIO"):
                self.listaProjBIO.append(projeto);

            if (projeto.tipo == "UFV"):
                self.listaProjUFV.append(projeto);

            if (projeto.tipo == "EOL"):
                self.listaProjEOL.append(projeto);

            if (projeto.tipo == "EOF"):
                self.listaProjEOF.append(projeto);

            if (projeto.tipo == "PCH"):
                self.listaProjPCH.append(projeto);

            return;
        
        def addProjetoRenovavelInt(self, projeto):
            self.listaProjRenovavelInt.append(projeto);
            return;
        
        def totalizaSeries(self):
            # inicializa em branco a matriz 
            self.hidroExTotal =[[0 for iper in range(self.numMesesTotal)] for icond in range(self.numCondicoes)];
            self.potDispExTotal =[[0 for iper in range(self.numMesesTotal)] for icond in range(self.numCondicoes)];
            self.ghMinTotal = 0
            
            # percorre series e periodos
            for usina in self.listaUHE:
                self.ghMinTotal = self.ghMinTotal + usina.ghMin;

                # percorre as linhas - numero de hidrologias
                for icond in range(0, self.numCondicoes):
                    # percorre as colunas - numero de meses total
                    for iper in range(0, self.numMesesTotal):
                        self.hidroExTotal[icond][iper] = self.hidroExTotal[icond][iper] + usina.serieHidrologica[icond][iper];
                        self.potDispExTotal[icond][iper] = self.potDispExTotal[icond][iper] + usina.potDisp[icond][iper];
                                
            # procura o valor minimo do vetor hidroExTotal
            self.ghMinTotalPer = [self.ghMinTotal for iper in range(0,self.numMesesTotal)];            
            for icond in range(0,self.numCondicoes):               
                for iper in range(0, self.numMesesTotal):
                    # pega o valor minimo
                    if self.hidroExTotal[icond][iper] < self.ghMinTotal:
                        self.ghMinTotalPer[iper] =self.hidroExTotal[icond][iper];

            return;