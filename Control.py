from RecebeDados import RecebeDados;
from Sistema import Sistema;
from Problema import Problema;
from SaidaDados import SaidaDados;
from SaidaNewave import SaidaNewave;
from contextlib import suppress;
from coopr.pyomo import *;
from pyomo.environ import *;
from pyomo.opt import *;
import os, jsonpickle;

class Control:
    
    def __init__(self, plan_dados, path, time):
        
        # carrega a planilha e recebe o caminho dela
        self.recebe_dados = RecebeDados(plan_dados);
        self.caminho = path;
        self.planilha = plan_dados;
        self.start = time;

        # carrega as configuracoes iniciais da planilha
        print("Carregando Dados");
        self.carregaInicio();

        # inicializa o sistema
        self.sin = Sistema(self.recebe_dados, self.tipoCombHidroEol);
        print("Carregando Problema");

        self.imprimeSeriesHidro();
 
        # cria o problema passando os parametros do carregaInicio
        self.problema = Problema(self.recebe_dados, self.sin, self.isRestPotHabilitada, self.isRestExpJanHabilitada, self.isPerpetHabilitada, self.fatorCarga, self.anoValidadeTemp, self.fatorValidadeTemp, self.isIntercambLimitado, self.subsFic);
        
        # habilita o cplex
        optsolver = SolverFactory("cplex", executable= "C:\\Program Files\\IBM\\ILOG\\CPLEX_Studio129\\cplex\\bin\\x64_win64\\cplex.exe");
        print ("Modelo Criado");
        self.problema.modelo.preprocess();
        print ("Pre-process executado");
        
        # configuracoes do solver
        optsolver.options['mipgap'] = 0.005;
        optsolver.options['mip_strategy_startalgorithm'] = 4;
        optsolver.options['lpmethod'] = 4;
        # congiguracao para manter reprodutibilidade entre casos 
        optsolver.options['parallel'] = 1;
        # configuracao tentar evitar erros de memoria
        optsolver.options['mip_strategy_file'] = 3;
        optsolver.options['emphasis_memory'] = 'y';
        optsolver.options['workmem'] = 12048;

        print("Executando o CPLEX");

        results = optsolver.solve(self.problema.modelo, load_solutions=True);#symbolic_solver_labels=True, tee=True);

        print("Impressão de Resultados");
        # escreve resultados em um txt
        with open(self.caminho + "resultado.txt", "w") as saidaResul:
            results.write(ostream=saidaResul);
            
        # inicializa o objeto de saida de dados
        self.saida_dados = SaidaDados(self.sin, self.problema, self.caminho, self.planilha, self.pastaCod, self.nomeSubs);

        # inicializa o objeto de saida para o newave
        self.saida_newave = SaidaNewave(self.recebe_dados, self.sin, self.problema, self.caminho, self.numSubs, self.subsNFic, self.subsFic, self.nomeSubs);
        
        # relaxa o problema
        self.problema.relaxar();

        # chama o metodo para limpar os duais da planilha
        self.saida_dados.limparDuais();
    
        # faz a preparacao e impressao dos duais escolhidos atraves da aba inicial na planilha
        # a letra passada como parametro eh um indicativo de qual dual deve ser impresso
        if self.isImpresso[0]:
            print("Resolvendo problema relaxado para Dual de Energia");
            self.problema.prepararDualEnergia();
            results = optsolver.solve(self.problema.modelo, load_solutions=True, warmstart=True);#, symbolic_solver_labels=True, tee=True);
            self.saida_dados.imprimeDuais("E");
        if self.isImpresso[1]:
            print("Resolvendo problema relaxado para Dual de Potencia");
            self.problema.prepararDualPotencia();
            results = optsolver.solve(self.problema.modelo, load_solutions=True, warmstart=True);#, symbolic_solver_labels=True, tee=True);
            self.saida_dados.imprimeDuais("P");
        if self.isImpresso[2]:
            print("Resolvendo problema relaxado para Dual Duplo");
            self.problema.prepararDualDuplo();
            results = optsolver.solve(self.problema.modelo, load_solutions=True, warmstart=True);#, symbolic_solver_labels=True, tee=True);
            self.saida_dados.imprimeDuais("D");
            
        self.saida_dados.imprimeLog(tempo_inicial = self.start);

        return;
    
    def carregaInicio(self):

        # pega o numero de subsistemas na planilha geral
        self.recebe_dados.defineAba("GERAL");
        self.nsis = int(self.recebe_dados.pegaEscalar("G10"));

        # verifica os parametros de impressao na aba inicial da planilha
        # tem que vir antes da criacao do problema pois na aba estao contidos os fatores de carga
        self.recebe_dados.defineAba("Inicial");

        # limpa o diretorio Temp para evitar problemas de memoria se o flag estiver ativo
        if (self.recebe_dados.pegaEscalar("O10")==1):
            print("Limpando arquivos temporarios")
            pastaTemp = "C:\\Users\\" + str(os.getlogin()) + "\\AppData\\Local\\Temp";
            for filename in os.listdir(pastaTemp):
                file_path = os.path.join(pastaTemp, filename);
                try:
                    if os.path.isfile(file_path) or os.path.islink(file_path):
                        os.unlink(file_path);
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path, ignore_errors=True);
                except:
                    print("Um ou mais arquivos temporarios nao puderam ser excluidos");
        
        # declara as variaveis que receberao a informacao das checkboxes e dos fatores na planilha
        self.isImpresso = [False for i in range(0,3)];
        self.isRestPotHabilitada = False;
        self.isExpJsonHabilitada = False;
        self.isRestExpJanHabilitada = False;
        self.isPerpetHabilitada = False;
        self.isIntercambLimitado = False;
        self.tipoCombHidroEol = "";
        self.fatorCarga = [];
        self.subsNFic = [];
        self.subsFic = [];
        self.nomeSubs = [];
        self.numSubs = [];
        self.anoValidadeTemp = 0;
        self.fatorValidadeTemp = [];

        #le a pasta na qual esta o codigo fonte
        self.pastaCod = str(self.recebe_dados.pegaEscalar("G7"));
        
        # faz a verificacao dos valores das celulas de fato. No excel, se for verdadeiro (checked), ele retorna o numero 1
        for i in range(0,3):
            if (self.recebe_dados.pegaEscalar("O4",lin_offset=i)==1):
                # se o valor da celula for true, a posicao da lista passa a ser true
                self.isImpresso[i] = True;
                
        # verifica se a restricao de potencia esta habilitada
        if (self.recebe_dados.pegaEscalar("O7")==1):
            # se o valor da celula for true, o parametro deve indicar que a restricao de potencia esta habilitada
            self.isRestPotHabilitada = True;
            
        # verifica se a opcao de exportar o objeto em json esta habilitada
        if (self.recebe_dados.pegaEscalar("O8")==1):
            # se o valor da celula for true, o parametro deve indicar que o json deve ser criado
            self.isExpJsonHabilitada = True;

        # verifica se a opcao de expansao apenas em janeiro esta habilitada
        if (self.recebe_dados.pegaEscalar("O3")==1):
            # se o valor da celula for true, o parametro deve indicar que a restricao de expansao somente em janeiro esta habilitada
            self.isRestExpJanHabilitada = True;

        # verifica se a opcao de perpetuidade esta habilitada
        if (self.recebe_dados.pegaEscalar("O2")==1):
            # se o valor da celula for true, o parametro deve indicar que o calculo da FO com perpetuidade esta habilitado
            self.isPerpetHabilitada = True;
        
        # importa da aba Inicial os valores dos fatores de carga dos subsistemas
        self.fatorCarga = self.recebe_dados.pegaVetor("A12", direcao="horizontal", tamanho=self.nsis, lin_offset=0, col_offset=0);

        # importa da aba Inicial os numeros subsistemas
        self.numSubs = self.recebe_dados.pegaVetor("A18", direcao="horizontal", tamanho=self.nsis, lin_offset=0, col_offset=0);

        # importa da aba Inicial os nomes dos subsistemas
        self.nomeSubs = self.recebe_dados.pegaVetor("A19", direcao="horizontal", tamanho=self.nsis, lin_offset=0, col_offset=0);

        # cria vetor de subs nao ficticios
        # se o numero do subs é maior que 99 é um subs ficticio
        for subs in range(self.nsis):
            if self.numSubs[subs] < 100:
                self.subsNFic.append(self.numSubs[subs]);
            else:
                self.subsFic.append(subs);

        # verifica se o usuario deseja usar todas as combinacoes de series hidrolicas e eolicas ou combinacoes intercaladas
        if (self.recebe_dados.pegaEscalar("O9")==1):
            self.tipoCombHidroEol = "completa";
        elif (self.recebe_dados.pegaEscalar("O10")==1):
            self.tipoCombHidroEol = "intercalada";
        else:
            print("Forma de incorporacao das series eolicas nao escolhida.");
        
        # verifica se o usuario deseja usar o limite de intercambio
        if (self.recebe_dados.pegaEscalar("O22")==1):
            self.isIntercambLimitado = True;

        # importa da aba inicial o ano de entrada da restricao de validade temporal
        self.anoValidadeTemp = self.recebe_dados.pegaEscalar("D22");
        
        # verifica se o usuario deseja usar a porcentagem de independencia da transmissao nos subsistemas ou nao
        if (self.recebe_dados.pegaEscalar("O23")==0):
            self.fatorValidadeTemp = [0 for isis in range(0, self.nsis)];
        else:
            self.fatorValidadeTemp = self.recebe_dados.pegaVetor("A25", direcao="horizontal", tamanho=self.nsis, lin_offset=0, col_offset=0);
        
        return;
    
    def exportaObjeto(self):
        # exporta o objeto do sistema do python pro Json
        objSistema = jsonpickle.encode(self.sin);
        saidaResul = open(self.caminho + "objetoSistema.json", "w");
        saidaResul.write(str(objSistema));
        
        # fecha o primeiro arquivo
        saidaResul.close();
        
        # exporta o objeto do modelo do python pro Json
        objModelo = jsonpickle.encode(self.problema);
        saidaResul = open(self.caminho + "objetoModelo.json", "w");
        saidaResul.write(str(objModelo));
        
        # fecha o segundo arquivo
        saidaResul.close();
        
        return;
    
    def importaObjeto(self, arqObjeto):
        # arqObjeto se refere ao arquivo que contem o objeto a ser importado do json pro python
        arquivo = open(self.caminho + arqObjeto);
        json_str = arquivo.read();
        restored_obj = jsonpickle.decode(json_str);
        list_objects = [restored_obj];
        print ("list_objects: ", list_objects);
   
        return;

    def imprimeSeriesHidro(self):
        sin = self.sin;
        
        # abre os arquivos
        saidaEner = open(self.caminho + "serieHidro.txt", "w");
        saidaPot = open(self.caminho + "pdispHidro.txt", "w");
        
        # percorre os cenarios
        for icen in range(sin.numHidros):
            
            # percorre primeiramente os projetos
            for isis in range(0,14):
                # imprime o nome da usina
                saidaEner.write(str(icen) + "," + str(isis));
                saidaPot.write(str(icen) + "," + str(isis));
                
                # percorre os periodos
                for iper in range(sin.numMeses): 
                    saidaEner.write("," + str(sin.subsistemas[isis].hidroExTotal[icen][iper]));
                    saidaPot.write("," + str(sin.subsistemas[isis].potDispExTotal[icen][iper]));
                
                # proxima linha
                saidaEner.write("\n");
                saidaPot.write("\n");
            
        # fecha o arquivo
        saidaEner.close();
        saidaPot.close();
        
        return;