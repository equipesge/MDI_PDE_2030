from RecebeDados import RecebeDados;
from Sistema import Sistema;
from UHE import UHE;
from Termica import Termica;
from Renovavel import Renovavel;
from Projeto import *;
from coopr.pyomo import *;

class Problema:
    
    def __init__(self, recebe_dados, sistema, habilitaRestPot, habilitaRestExpJan, habilitaPerpet, vetFatorCarga, anoValidade, vetFatorValidade, intercLimitado, subsFic):
        # recebe como parametro o sistema em que estao as informacoes incorporadas do excel
        self.fonte_dados = recebe_dados;
        self.sin = sistema;
        self.isRestPotHabilitada = habilitaRestPot;
        self.isRestExpJanHabilitada = habilitaRestExpJan;
        self.isPerpetHabilitada = habilitaPerpet;
        self.fatorCarga = vetFatorCarga;
        self.fatorValidadeTemp = vetFatorValidade;
        self.isIntercambLimitado = intercLimitado;
        self.subsFic = subsFic;

        # calcula o periodo de entrada da restricao
        self.perValidadeTemp = (anoValidade - self.sin.anoInicial)*12;

        # monta o problema com os dados do sistema
        self.montaProblema();
       
        return;
    
    def montaProblema (self):
        # cria um Concrete Model do Pyomo
        self.modelo = ConcreteModel();
        
        # atribui o objeto sin ao modelo (pois utiliaremos nas funcoes do pyomo)
        self.modelo.sin = self.sin;        
        
        # cria os conjuntos basicos
        self.criaConjuntosBasicos();
        
        # cria as variaveis de decisao
        self.criaVariaveisDecisao();
    
        # cria os parametros
        self.criaParametros();

        # cria as restricoes
        self.criaRestricoes();
        
        # cria a funcao objetivo
        self.objetivo();

        return;
        
    def criaConjuntosBasicos(self):
        modelo = self.modelo;
        
        # conjunto de sistemas
        modelo.subsistemas = Set(initialize=range(0, self.sin.nsis));
        
        # conjunto de patamares
        modelo.patamares = Set(initialize=range(0, self.sin.nPatamares));

        # conjunto de uhes existentes
        modelo.uheExist = Set(initialize=[usina.nomeUsina for isis in modelo.subsistemas for usina in self.sin.subsistemas[isis].listaUHE]);
        
        # conjunto de uhes novas - projetos
        modelo.projUHENova = Set(initialize=[usina.nomeUsina for isis in modelo.subsistemas for usina in self.sin.subsistemas[isis].listaProjUHE]);
        
        # conjunto de termicas existentes        
        modelo.termExist = Set(initialize=[usina.nomeUsina for isis in modelo.subsistemas for usina in self.sin.subsistemas[isis].listaTermica]);
        
        # conjunto de termicas novas continuas - projetos
        modelo.projTermCont = Set(initialize=[usina.nomeUsina for isis in modelo.subsistemas for usina in self.sin.subsistemas[isis].listaProjTermica]);
        
        # conjunto de projetos renovaveis continuos
        modelo.projRenovCont = Set(initialize=[usina.nomeUsina for isis in modelo.subsistemas for usina in self.sin.subsistemas[isis].listaProjRenovavel]);

        # conjunto de projetos de reversiveis
        modelo.projReversivel = Set(initialize=[usina.nomeUsina for isis in modelo.subsistemas for usina in self.sin.subsistemas[isis].listaProjReversivel]);
        
        # conjunto de periodos e anos
        modelo.periodos = Set(initialize=range(0,self.sin.numMeses));
        modelo.periodosTotal = Set(initialize=range(0,self.sin.numMesesTotal));
        modelo.anos = Set(initialize=range(0,int(self.sin.numMesesTotal/12)));

        # conjunto de condicoes
        modelo.condicoes = Set(initialize=range(0,self.sin.numCondicoes));

        # conjunto de restricoes adicionais
        modelo.conjSteps = Set(initialize=range(len(self.sin.restricoes.Step)));
        modelo.conjLimiteAno = Set(initialize=range(len(self.sin.restricoes.LimiteAno)));
        modelo.conjIgualdade = Set(initialize=range(len(self.sin.restricoes.Igualdade)));
        modelo.conjIgualdadeMax = Set(initialize=range(len(self.sin.restricoes.IgualdadeMax)));
        modelo.conjLimiteIncAno = Set(initialize=range(len(self.sin.restricoes.LimiteIncAno)));
        modelo.conjProporcao = Set(initialize=range(len(self.sin.restricoes.Proporcao)));
        modelo.conjAgrint = Set(initialize=range(len(self.sin.agrints)));
        
        return;

    # regra para verificar se a termica é binaria ou continua
    def verificaTermBinaria(self, modelo, projInd, perInd):
        # recupera o projeto na lista de projetos utilizando o indice passado para a funcao
        proj = modelo.sin.listaGeralProjTerm[projInd];
        # verifica se o projeto eh continua ou nao e retorna o dominio da variavel
        if (proj.isContinua):
            return (NonNegativeReals);
        else:
            return (Binary);
    
    def criaVariaveisDecisao(self):
        modelo = self.modelo;
        
        # variaveis de operacao (energia)
        modelo.prodHidroExist = Var(modelo.subsistemas, modelo.patamares, modelo.periodosTotal, modelo.condicoes, domain=NonNegativeReals);
        modelo.prodHidroNova = Var(modelo.projUHENova, modelo.patamares, modelo.periodosTotal, modelo.condicoes, domain=NonNegativeReals);
        modelo.prodTerm = Var(modelo.termExist, modelo.patamares, modelo.periodosTotal, modelo.condicoes, domain=NonNegativeReals);         
        modelo.deficit = Var(modelo.subsistemas, modelo.patamares, modelo.periodosTotal, modelo.condicoes, domain=NonNegativeReals);
        modelo.interc = Var(modelo.subsistemas, modelo.subsistemas, modelo.patamares, modelo.periodosTotal, modelo.condicoes, domain=NonNegativeReals);
        modelo.prodReversivel = Var(modelo.projReversivel, modelo.patamares, modelo.periodosTotal, modelo.condicoes, domain=NonNegativeReals);
        modelo.bombReversivel = Var(modelo.projReversivel, modelo.patamares, modelo.periodosTotal, modelo.condicoes, domain=NonNegativeReals);
                
        # variaveis de expansao
        modelo.investHidro = Var(modelo.projUHENova, modelo.periodosTotal, domain=Binary);              
        modelo.capTermCont = Var(modelo.projTermCont, modelo.periodosTotal, domain=self.verificaTermBinaria);
        modelo.capRenovCont = Var(modelo.projRenovCont, modelo.periodosTotal, domain=NonNegativeReals);
        modelo.capExpInter = Var(modelo.subsistemas, modelo.subsistemas, modelo.periodosTotal, domain=NonNegativeReals);
        modelo.prodTermCont = Var(modelo.projTermCont, modelo.patamares, modelo.periodosTotal, modelo.condicoes, domain=NonNegativeReals);
        modelo.capReversivel = Var(modelo.projReversivel, modelo.periodosTotal, domain=NonNegativeReals);
        
        # variaveis para demanda de potencia
        modelo.intercPot = Var(modelo.subsistemas, modelo.subsistemas, modelo.periodosTotal, modelo.condicoes, domain=NonNegativeReals);
        modelo.deficitPot = Var(modelo.subsistemas, modelo.periodosTotal, modelo.condicoes, domain=NonNegativeReals);
        
        # penalidade
        modelo.penalidadeGHMinExist = Var(modelo.subsistemas, modelo.patamares, modelo.periodosTotal, modelo.condicoes, domain=NonNegativeReals);
        modelo.penalidadeGHMinNova = Var(modelo.projUHENova, modelo.patamares, modelo.periodosTotal, modelo.condicoes, domain=NonNegativeReals);

        # varaveis de step
        modelo.steps = Var(modelo.conjSteps, domain=NonNegativeReals);
        
        # variaveis relacionadas ao CME de Energia, Potencia e Duplo
        modelo.dE = Var(modelo.anos, domain=NonNegativeReals);
        modelo.dP = Var(modelo.anos, domain=NonNegativeReals);
        modelo.dD = Var(modelo.anos, domain=NonNegativeReals);
        
        # variaveis necessarias para representar a motorizacao de projetos de UHE
        modelo.moto = Var(modelo.projUHENova, modelo.periodosTotal, domain=NonNegativeReals);
        
        return;
    
    def criaParametros(self):
        modelo = self.modelo;
        
        # hashs para a correta atribuicao dos valores de custo de operacao e capacidade
        v={};v_inflex={}
        for sis in self.sin.subsistemas: 
            for usina in sis.listaTermica:
                for iper in modelo.periodosTotal:
                    v[usina.nomeUsina, iper] = usina.cvu[iper];
                    v_inflex[usina.nomeUsina, iper] = usina.inflexExistente[iper%12];              
       
        # custos de operacao
        modelo.cvuTermExist = Param(modelo.termExist, modelo.periodosTotal, initialize=v);
        
        # inflexibilidades
        modelo.inflexTermica = Param(modelo.termExist, modelo.periodosTotal, initialize=v_inflex);
        
        # hash para a correta atribuicao do custo de deficit no modelo
        v={};      
        for isis in modelo.subsistemas:
            for ipat in modelo.patamares: 
                v[isis, ipat] = self.sin.custoDefc[ipat]; 
        # custo de deficit                     
        modelo.custoDefc = Param(modelo.subsistemas, modelo.patamares, initialize=v);
        
        # hashs para atribuicao dos recursos e requisitos. Variam para cada subsistema, cada periodo de tempo e cada cenario hidrologico
        v={};v_h={};v_r_PCH={};v_r_EOL={};v_r_UFV={};v_r_BIO={};v_p={};    
        for isis in modelo.subsistemas: 
            for iper in modelo.periodosTotal:
                for icen in modelo.condicoes:
                    v_h[isis,iper,icen] = self.sin.subsistemas[isis].hidroExTotal[icen][iper];
                    v_p[isis,iper,icen] = self.sin.subsistemas[isis].potDispExTotal[icen][iper];
                    v_r_PCH[isis,iper,icen] = self.sin.subsistemas[isis].montanteRenovExPCH[iper];
                    v_r_EOL[isis,iper,icen] = self.sin.subsistemas[isis].montanteRenovExEOL[iper];
                    v_r_UFV[isis,iper,icen] = self.sin.subsistemas[isis].montanteRenovExUFV[iper];
                    v_r_BIO[isis,iper,icen] = self.sin.subsistemas[isis].montanteRenovExBIO[iper];
        
        # considerando que a demanda varia com o patamar            
        for isis in modelo.subsistemas:
            for ipat in modelo.patamares: 
                for iper in modelo.periodosTotal:
                    for icen in modelo.condicoes:
                        v[isis,ipat,iper,icen] = self.sin.subsistemas[isis].demandaEnerg[ipat][iper];
        
        # recursos e requisitos                    
        modelo.demanda = Param(modelo.subsistemas, modelo.patamares, modelo.periodosTotal, modelo.condicoes, initialize=v);
        modelo.energiaHidroEx = Param(modelo.subsistemas, modelo.periodosTotal, modelo.condicoes, initialize=v_h);
        modelo.pDispHidroEx = Param(modelo.subsistemas, modelo.periodosTotal, modelo.condicoes, initialize=v_p);

        # energia renovavel
        modelo.enPCHEx = Param(modelo.subsistemas, modelo.periodosTotal, modelo.condicoes, initialize=v_r_PCH);
        modelo.enEOLEx = Param(modelo.subsistemas, modelo.periodosTotal, modelo.condicoes, initialize=v_r_EOL);
        modelo.enUFVEx = Param(modelo.subsistemas, modelo.periodosTotal, modelo.condicoes, initialize=v_r_UFV);
        modelo.enBIOEx = Param(modelo.subsistemas, modelo.periodosTotal, modelo.condicoes, initialize=v_r_BIO);

        # custos de investimento em hidro, renovavel continua e termi continua
        cinv_h={};cinv_rc={};cinv_ri={};en={};cinv_tc={};cvu_tc={};fdisp_tc={};cinv_rev={};rend_rev={};dataMin_rev={};
        for sis in self.sin.subsistemas: 
            # percorre as usinas
            for usina in sis.listaProjUHE:
                # preenche investimento de uhe 
                cinv_h[usina.nomeUsina] = usina.custoFixo;
                for iper in modelo.periodosTotal:
                    for icen in modelo.condicoes:
                        # preenche a energia da UHE nova
                        en[usina.nomeUsina,iper,icen] = usina.serieHidrologica[icen][iper];

            # preenche o custo de investimento das renovaveis
            for usina in sis.listaProjRenovavel:
                cinv_rc[usina.nomeUsina] = usina.custoMensal;

            # preenche o custo de investimento das reversiveis
            for usina in sis.listaProjReversivel:
                cinv_rev[usina.nomeUsina] = usina.custoMensal;
                rend_rev[usina.nomeUsina] = usina.rendimento;
                dataMin_rev[usina.nomeUsina] = usina.dataMinima;
            
            # preenche o custo de investimento das termicas    
            for term in sis.listaProjTermica:
                for iper in modelo.periodosTotal:
                    cvu_tc[term.nomeUsina, iper] = term.cvu[iper]; 
                cinv_tc[term.nomeUsina] = term.custoFixo;
                fdisp_tc[term.nomeUsina] = term.fdisp;

        # armazena nos params do modelo de investimento
        modelo.custoInvHidro = Param(modelo.projUHENova, initialize=cinv_h);
        modelo.custoInvRenovCont = Param(modelo.projRenovCont, initialize=cinv_rc);
        modelo.energHidroNova = Param(modelo.projUHENova, modelo.periodosTotal, modelo.condicoes, initialize=en);
        modelo.cvuProjTerm = Param(modelo.projTermCont, modelo.periodosTotal, initialize=cvu_tc);
        modelo.custoInvProjTerm = Param(modelo.projTermCont, initialize=cinv_tc);
        modelo.fdispProjTerm = Param(modelo.projTermCont, initialize=fdisp_tc);
        modelo.custoInvReversivel = Param(modelo.projReversivel, initialize=cinv_rev);

        # rendimento reversíveis
        modelo.rendReversivel = Param(modelo.projReversivel, initialize=rend_rev);

        # data minima reversíveis
        modelo.dataMinReversivel = Param(modelo.projReversivel, initialize=dataMin_rev);
        
        # parametros correspondentes ao valores para habilitar os duais
        modelo.rhsDD = Param(mutable=True, default=0);
        modelo.rhsDE = Param(mutable=True, default=0.3);
        modelo.rhsDP = Param(mutable=True, default=0);
        
        return;
    
    def criaRestricoes(self):
        
        self.criaRestricoesDemanda();
        self.criaRestricoesCapacidades();
        self.criaRestricoesInvestimentos();
        self.criaRestricoesFisicas();
        self.criaRestricoesAdicionais();
        self.criaRestricoesDuais();
        self.criaRestricoesHidro();
        self.criaRestricoesReversiveis();
        if self.isRestExpJanHabilitada:  # se o flag de expansao apenas em janeiro/julho estiver habilitado, cria as restriçoes
            self.criaRestricoesReducaoProblema();
        if self.isPerpetHabilitada:  # se o flag de perpetuidade estiver habilitado, cria as variaveis incrementais
            self.createIncrementais();
        return;

    def criaRestricoesHidro(self):
        modelo = self.modelo;
        modelo.sin = self.sin;
        
        # produto da producao e patamar eh igual a serie de energia 
        def resSomatorioHidroNovas(modelo, iuhe, iper, icen):
            return ( sum(modelo.prodHidroNova[iuhe, ipat, iper, icen]*modelo.sin.duracaoPatamar[ipat][iper] for ipat in modelo.patamares) <= ( modelo.energHidroNova[iuhe,iper,icen]*(sum(modelo.investHidro[iuhe,iperaux] for iperaux in range(0,iper+1))) ) );
        modelo.somatorioHidroNovas = Constraint(modelo.projUHENova, modelo.periodosTotal, modelo.condicoes, rule=resSomatorioHidroNovas);
        
        # restricoes que acoplam a motorizacao das usinas novas com seu respectivo investimento
        def resMotorizacao(modelo, iuhe, iper):
            return ( sum(modelo.moto[iuhe, tau] for tau in range(iper+1)) <= sum(modelo.investHidro[iuhe,tau2]  for tau2 in range(iper+1)) );
        modelo.acoplaMotorizacao = Constraint(modelo.projUHENova, modelo.periodosTotal, rule=resMotorizacao);
        
        # restricoes que limitam a taxa de motorizacao
        def resTaxaMotorizacao(modelo, iuhe, iper):
            # pega a uhe e calcula a taxa de motorizacao
            proj = modelo.sin.listaGeralProjUHE[iuhe];
            r=1;
            if proj.nMesesMotorizacao > 0:
                r = 1/proj.nMesesMotorizacao;
            return modelo.moto[iuhe,iper] <= r;        
        modelo.taxaMotorizacao = Constraint(modelo.projUHENova, modelo.periodosTotal, rule=resTaxaMotorizacao);
        
        # GHmax das hidro novas - restrito pela motorizacao
        def restGHMaxHidroNovas(modelo, iuhe, ipat, iper, icen):
            return (modelo.prodHidroNova[iuhe, ipat, iper, icen] <= modelo.sin.listaGeralProjUHE[iuhe].potDisp[icen][iper] * sum(modelo.moto[iuhe, tau] for tau in range(iper+1)));
        modelo.GHMaxHidroNovas = Constraint(modelo.projUHENova, modelo.patamares, modelo.periodosTotal, modelo.condicoes, rule=restGHMaxHidroNovas);
        
        # GHmin das hidro novas - aplica a motorizacao para reduzir o ghmin
        def restGHMinHidroNovas(modelo, iuhe, ipat, iper, icen):
            return (modelo.prodHidroNova[iuhe, ipat, iper, icen] + modelo.penalidadeGHMinNova[iuhe, ipat, iper, icen] >= modelo.sin.listaGeralProjUHE[iuhe].ghMin * sum(modelo.moto[iuhe, tau] for tau in range(iper+1)) );
        modelo.GHMinHidroNovas = Constraint(modelo.projUHENova, modelo.patamares, modelo.periodosTotal, modelo.condicoes, rule=restGHMinHidroNovas);
        
        # Soma produto hidro existentes
        def restSomatorioHidroExist(modelo, isis, iper, icen):
            return (sum(modelo.prodHidroExist[isis, ipat, iper, icen]*modelo.sin.duracaoPatamar[ipat][iper] for ipat in modelo.patamares) <= modelo.energiaHidroEx[isis, iper, icen]);
        modelo.somatorioHidroExist = Constraint(modelo.subsistemas, modelo.periodosTotal, modelo.condicoes, rule=restSomatorioHidroExist);

        # GHmax das hidro existentes
        def restGHMaxHidroExist(modelo, isis, ipat, iper, icen):
            return (modelo.prodHidroExist[isis, ipat, iper, icen] <= modelo.pDispHidroEx[isis,iper,icen]);
        modelo.GHMaxHidroExist= Constraint(modelo.subsistemas, modelo.patamares, modelo.periodosTotal, modelo.condicoes, rule=restGHMaxHidroExist);
        
        # GHmin das hidro existentes
        def restGHMinHidroExist(modelo, isis, ipat, iper, icen):
            return (modelo.prodHidroExist[isis, ipat, iper, icen] + modelo.penalidadeGHMinExist[isis, ipat, iper, icen] >= modelo.sin.subsistemas[isis].ghMinTotalPer[iper]);
        modelo.GHMinHidroExist = Constraint(modelo.subsistemas, modelo.patamares, modelo.periodosTotal, modelo.condicoes, rule=restGHMinHidroExist);

        return;

    def criaRestricoesReversiveis(self):
        modelo = self.modelo;
        modelo.sin = self.sin;

        # restricao de investimento crescente reversivel
        def restNaoInvReversivel (modelo,iproj,iper):
            if iper < modelo.dataMinReversivel[iproj]:
                return (modelo.capReversivel[iproj,iper] == 0); 
            else:
                return (modelo.capReversivel[iproj,iper] >= modelo.capReversivel[iproj,iper-1]); 
        modelo.naoInvReversivel = Constraint(modelo.projReversivel, modelo.periodosTotal, rule=restNaoInvReversivel);

        # geracao maxima reversivel
        def resProdMaxReversivel(modelo, iproj, iper, icen, ipat):
            return (modelo.prodReversivel[iproj, ipat, iper, icen] <= modelo.capReversivel[iproj,iper]);
        modelo.prodMaxReversivel = Constraint(modelo.projReversivel, modelo.periodosTotal, modelo.condicoes, modelo.patamares, rule=resProdMaxReversivel);

        # bombeamento maximo reversivel
        def resBombMaxReversivel(modelo, iproj, iper, icen, ipat):
            return (modelo.bombReversivel[iproj, ipat, iper, icen] <= modelo.capReversivel[iproj,iper]);
        modelo.bombMaxReversivel = Constraint(modelo.projReversivel, modelo.periodosTotal, modelo.condicoes, modelo.patamares, rule=resBombMaxReversivel);

        # somatorio patamares reversivel
        def resSomaPatReversivel(modelo, iproj, iper, icen):
            return (sum(modelo.prodReversivel[iproj, ipat, iper, icen]*modelo.sin.duracaoPatamar[ipat][iper] for ipat in modelo.patamares) <= modelo.rendReversivel[iproj]*sum(modelo.bombReversivel[iproj, ipat, iper, icen]*modelo.sin.duracaoPatamar[ipat][iper] for ipat in modelo.patamares));
        modelo.somaPatReversivel = Constraint(modelo.projReversivel, modelo.periodosTotal, modelo.condicoes, rule=resSomaPatReversivel);

        # restringe geracao reversivel - pode gerar na ponta, pesada e media e bombear na leve e media
        def resPatProdReversivel(modelo, iproj, ipat, iper, icen):
            return (modelo.prodReversivel[iproj, ipat, iper, icen] == 0);
        modelo.patProdReversivel = Constraint(modelo.projReversivel, modelo.sin.naoGeraReversivel, modelo.periodosTotal, modelo.condicoes, rule=resPatProdReversivel);

        # restringe bombeamento reversivel - pode gerar na ponta, pesada e media e bombear na leve e media
        def resPatBombReversivel(modelo, iproj, ipat, iper, icen):
            return (modelo.bombReversivel[iproj, ipat, iper, icen] == 0);
        modelo.patBombReversivel = Constraint(modelo.projReversivel, modelo.sin.naoBombReversivel, modelo.periodosTotal, modelo.condicoes, rule=resPatBombReversivel);

        return;

    def criaRestricoesAdicionais(self):
        modelo = self.modelo;
        modelo.sin = self.sin;

        # restricao de step
        def resStepUB(modelo, iStep, iper):
            # a restricao subtrai os ultimos 12 meses
            step = modelo.sin.restricoes.Step[iStep];
            if ((iper >= step.anoInicial*12) and (iper < (step.anoFinal+1)*12) and (iper%12 == step.mes)): #restricao de step valida apenas para os meses de dezembro
                if iper >= 12:
                    if (step.tipoProj == "RenovCont"):
                        return (sum(modelo.capRenovCont[renov,iper] for renov in modelo.sin.restricoes.Step[iStep].listaProj) - sum(modelo.capRenovCont[renov,iper-12] for renov in modelo.sin.restricoes.Step[iStep].listaProj) == modelo.steps[iStep]);
                    elif (step.tipoProj == "Term"):
                        return (sum(modelo.capTermCont[term,iper]*modelo.sin.listaGeralProjTerm[term].potUsina for term in modelo.sin.restricoes.Step[iStep].listaProj) - sum(modelo.capTermCont[term,iper-12] for term in modelo.sin.restricoes.Step[iStep].listaProj) == modelo.steps[iStep]);
                # se o periodo estiver no primeiro ano fica restrito ao proprio step
                else:
                    if (step.tipoProj == "RenovCont"):
                        return (sum(modelo.capRenovCont[renov,iper] for renov in modelo.sin.restricoes.Step[iStep].listaProj) == modelo.steps[iStep]);
                    elif (step.tipoProj == "Term"):
                        return (sum(modelo.capTermCont[term,iper]*modelo.sin.listaGeralProjTerm[term].potUsina for term in modelo.sin.restricoes.Step[iStep].listaProj) == modelo.steps[iStep]);
            else: return Constraint.Feasible;

        modelo.atendeStepUB = Constraint(modelo.conjSteps, modelo.periodosTotal, rule=resStepUB);

        # minimos e maximos de steps
        def restStepMin(modelo, iStep):
            return modelo.steps[iStep] >= modelo.sin.restricoes.Step[iStep].val_min;
        def restStepMax(modelo, iStep):
            return modelo.steps[iStep] <= modelo.sin.restricoes.Step[iStep].val_max;
        modelo.stepMin = Constraint(modelo.conjSteps, rule=restStepMin);
        modelo.stepMax = Constraint(modelo.conjSteps, rule=restStepMax);

        # restricao de limite
        def resLimiteAno(modelo, iRest, iper):
            # a restricao eh valida apenas para o periodo definido
            rest = modelo.sin.restricoes.LimiteAno[iRest];
            if ((iper >= rest.anoInicial*12) and (iper < (rest.anoFinal+1)*12) and ((iper%12) == rest.mes)):
                if (rest.tipoProj == "RenovCont"):
                    return (sum(modelo.capRenovCont[renov,iper] for renov in modelo.sin.restricoes.LimiteAno[iRest].listaProj) <= modelo.sin.restricoes.LimiteAno[iRest].valores[int(iper/12)]);
                if (rest.tipoProj == "Reversivel"):
                    return (sum(modelo.capReversivel[reversivel,iper] for reversivel in modelo.sin.restricoes.LimiteAno[iRest].listaProj) <= modelo.sin.restricoes.LimiteAno[iRest].valores[int(iper/12)]);
                if (rest.tipoProj == "Term"):
                    return (sum(modelo.capTermCont[term,iper]*modelo.sin.listaGeralProjTerm[term].potUsina for term in modelo.sin.restricoes.LimiteAno[iRest].listaProj) <= modelo.sin.restricoes.LimiteAno[iRest].valores[int(iper/12)]);
            else: 
                return Constraint.Feasible;
        modelo.LimiteAno = Constraint(modelo.conjLimiteAno, modelo.periodosTotal, rule=resLimiteAno);

        # restricao de limite incremental
        def resLimiteIncAno(modelo, iRest, iper):
            # a restricao eh valida apenas para o periodo definido
            rest = modelo.sin.restricoes.LimiteIncAno[iRest];
            if ((iper >= rest.anoInicial*12) and (iper < (rest.anoFinal+1)*12)):
                # o retorno da restricao depende do tipo de projeto
                if (rest.tipoProj == "RenovCont"):
                    # verifica se esta no primeiro ano da restricao
                    if ((iper - rest.anoInicial*12)>=12):
                        return (sum((modelo.capRenovCont[renov,iper]-modelo.capRenovCont[renov,iper-12]) for renov in modelo.sin.restricoes.LimiteIncAno[iRest].listaProj) <= modelo.sin.restricoes.LimiteIncAno[iRest].valores[int(iper/12)]);
                    else:
                        return (sum((modelo.capRenovCont[renov,iper]) for renov in modelo.sin.restricoes.LimiteIncAno[iRest].listaProj) <= modelo.sin.restricoes.LimiteIncAno[iRest].valores[int(iper/12)]);
                if (rest.tipoProj == "Term"):
                    # verifica se esta no primeiro ano da restricao
                    if ((iper - rest.anoInicial*12)>=12):
                        return (sum((modelo.capTermCont[term,iper]*modelo.sin.listaGeralProjTerm[term].potUsina-modelo.capTermCont[term,iper-12]*modelo.sin.listaGeralProjTerm[term].potUsina) for term in modelo.sin.restricoes.LimiteIncAno[iRest].listaProj) <= modelo.sin.restricoes.LimiteIncAno[iRest].valores[int(iper/12)]);
                    else:
                        return (sum((modelo.capTermCont[term,iper]) for term in modelo.sin.restricoes.LimiteIncAno[iRest].listaProj) <= modelo.sin.restricoes.LimiteIncAno[iRest].valores[int(iper/12)]);
                if (rest.tipoProj == "Reversivel"):
                    # verifica se esta no primeiro ano da restricao
                    if ((iper - rest.anoInicial*12)>=12):
                        return (sum((modelo.capReversivel[term,iper]-modelo.capReversivel[term,iper-12]) for term in modelo.sin.restricoes.LimiteIncAno[iRest].listaProj) <= modelo.sin.restricoes.LimiteIncAno[iRest].valores[int(iper/12)]);
                    else:
                        return (sum((modelo.capReversivel[term,iper]) for term in modelo.sin.restricoes.LimiteIncAno[iRest].listaProj) <= modelo.sin.restricoes.LimiteIncAno[iRest].valores[int(iper/12)]);
            else: 
                return Constraint.Feasible;
        modelo.LimiteIncAno = Constraint(modelo.conjLimiteIncAno, modelo.periodosTotal, rule=resLimiteIncAno);

        # restricao de igualdade
        def resIgualdade(modelo, iRest, iper):
            # a restricao eh valida apenas para o periodo definido
            rest = modelo.sin.restricoes.Igualdade[iRest];
            
            # o caso de hidro eh informado o mes
            if (rest.tipoProj == "Hidro"):
                # se o periodo for posterior ao do inicial, seta igual a 1 para todos os projetos da lista
                if (iper == (rest.anoInicial*12 + rest.mes)):
                    return (sum(modelo.investHidro[hidro,iper] for hidro in rest.listaProj) == len(rest.listaProj));
                else:
                    # caso contrario deve ser nulo
                    return Constraint.Feasible;
            
            # demais tipos de projetos vale o intervalo definido da restricao
            elif ((iper >= rest.anoInicial*12) and (iper < (rest.anoFinal+1)*12)):
                if (rest.tipoProj == "Term"):
                    return (sum(modelo.capTermCont[term,iper]*modelo.sin.listaGeralProjTerm[term].potUsina for term in modelo.sin.restricoes.Igualdade[iRest].listaProj) == modelo.sin.restricoes.Igualdade[iRest].valores[int(iper/12)]);
                if (rest.tipoProj == "RenovCont"):
                    return (sum(modelo.capRenovCont[renov,iper] for renov in modelo.sin.restricoes.Igualdade[iRest].listaProj) == modelo.sin.restricoes.Igualdade[iRest].valores[int(iper/12)]);
                if (rest.tipoProj == "Reversivel"):
                    return (sum(modelo.capReversivel[rev,iper] for rev in modelo.sin.restricoes.Igualdade[iRest].listaProj) == modelo.sin.restricoes.Igualdade[iRest].valores[int(iper/12)]);
            else: 
                return Constraint.Feasible;
        modelo.Igualdades = Constraint(modelo.conjIgualdade, modelo.periodosTotal, rule=resIgualdade);

        # restricao de igualdade maxima
        def resIgualdadeMax(modelo, iRest):
            # a restricao eh valida apenas para o periodo definido
            rest = modelo.sin.restricoes.IgualdadeMax[iRest];
            
            # o caso de hidro eh informado o mes
            if (rest.tipoProj == "Hidro"):
                # somatorio da variavel investHidro tem que ser igual a 1 ate o periodo de igualdade maxima
                return (sum(modelo.investHidro[hidro,iper] for hidro in rest.listaProj for iper in range(rest.anoInicial*12 + rest.mes)) == len(rest.listaProj));
            else: 
                return Constraint.Feasible;

        modelo.IgualdadesMax = Constraint(modelo.conjIgualdadeMax, rule=resIgualdadeMax);
        

        # restricao de proporcao entre duas fontes
        def resProporcao(modelo, iRest, iper):
            # a restricao eh valida apenas para o periodo definido
            rest = modelo.sin.restricoes.Proporcao[iRest];
            iper_rest = int((iper - rest.anoInicial*12)/12);
            if ((iper >= rest.anoInicial*12) and (iper < (rest.anoFinal+1)*12)):
                if (rest.tipoProj == "Term"):
                    if ((iper - rest.anoInicial*12)>=12):
                        return ((modelo.capTermCont[modelo.sin.restricoes.Proporcao[iRest].listaProj[0],iper]-modelo.capTermCont[modelo.sin.restricoes.Proporcao[iRest].listaProj[0],iper-12]) \
                                == (sum(modelo.capTermCont[term,iper]-modelo.capTermCont[term,iper-12] for term in modelo.sin.restricoes.Proporcao[iRest].listaProj)* \
                                (modelo.sin.restricoes.Proporcao[iRest].valoresProj[0][iper_rest])/(sum(modelo.sin.restricoes.Proporcao[iRest].valoresProj[term][iper_rest] for term in range(len(modelo.sin.restricoes.Proporcao[iRest].listaProj))))));
                    else:
                        return (modelo.capTermCont[modelo.sin.restricoes.Proporcao[iRest].listaProj[0],iper] \
                                == (sum(modelo.capTermCont[term,iper] for term in modelo.sin.restricoes.Proporcao[iRest].listaProj)*\
                                (modelo.sin.restricoes.Proporcao[iRest].valoresProj[0][iper_rest])/(sum(modelo.sin.restricoes.Proporcao[iRest].valoresProj[term][iper_rest] for term in range(len(modelo.sin.restricoes.Proporcao[iRest].listaProj))))));
                if (rest.tipoProj == "RenovCont"):
                    if ((iper - rest.anoInicial*12)>=12):
                        return ((modelo.capRenovCont[modelo.sin.restricoes.Proporcao[iRest].listaProj[0],iper]-modelo.capRenovCont[modelo.sin.restricoes.Proporcao[iRest].listaProj[0],iper-12]) \
                                == (sum(modelo.capRenovCont[renov,iper]-modelo.capRenovCont[renov,iper-12] for renov in modelo.sin.restricoes.Proporcao[iRest].listaProj)* \
                                (modelo.sin.restricoes.Proporcao[iRest].valoresProj[0][iper_rest])/(sum(modelo.sin.restricoes.Proporcao[iRest].valoresProj[renov][iper_rest] for renov in range(len(modelo.sin.restricoes.Proporcao[iRest].listaProj))))));
                    else:
                        return (modelo.capRenovCont[modelo.sin.restricoes.Proporcao[iRest].listaProj[0],iper] \
                                == (sum(modelo.capRenovCont[renov,iper] for renov in modelo.sin.restricoes.Proporcao[iRest].listaProj)*\
                                (modelo.sin.restricoes.Proporcao[iRest].valoresProj[0][iper_rest])/(sum(modelo.sin.restricoes.Proporcao[iRest].valoresProj[renov][iper_rest] for renov in range(len(modelo.sin.restricoes.Proporcao[iRest].listaProj))))));
            else: 
                return Constraint.Feasible;
        modelo.Proporcao = Constraint(modelo.conjProporcao, modelo.periodosTotal, rule=resProporcao);

        return;

    def criaRestricoesDemanda(self):
        modelo = self.modelo;
        modelo.sin = self.sin;
        
        # Restricao de atendimento a demanda de energia em cada cenario
        def restAtendeDemanda (modelo, isis, ipat, iper, icen):
            subsis = modelo.sin.subsistemas[isis];
            subs = self.sin.subsistemas[isis];
            if (self.sin.tipoCombHidroEol == "completa"):
                icond = icen;
            elif (self.sin.tipoCombHidroEol == "intercalada"):
                icond = icen % self.sin.numEol;
            else:
                print("opcao de combinacao de series hidrologicas com eolicas nao marcada");
            # atendimento a demanda de energia
            return (modelo.prodHidroExist[isis,ipat,iper,icen] \
                    + sum(modelo.prodTerm[term.nomeUsina,ipat,iper,icen] for term in subsis.listaTermica) \
                    + modelo.enPCHEx[isis,iper,icen] * subs.fatorPatPCH[ipat][iper%12] \
					+ modelo.enEOLEx[isis,iper,icen] * subs.fatorPatEOLEx[ipat][iper%12] \
					+ modelo.enUFVEx[isis,iper,icen] * subs.fatorPatUFVEx[ipat][iper%12] \
					+ modelo.enBIOEx[isis,iper,icen] * subs.fatorPatBIO[ipat][iper%12] \
                    + sum((1-subsis.perdasInterc[jsis][ipat][iper])*modelo.interc[jsis,isis,ipat,iper,icen] for jsis in modelo.subsistemas) - sum(modelo.interc[isis,jsis,ipat,iper,icen] for jsis in modelo.subsistemas) \
                    + sum(modelo.prodHidroNova[proj.nomeUsina,ipat,iper,icen] for proj in subsis.listaProjUHE) \
                    + sum(modelo.prodTermCont[proj.nomeUsina,ipat,iper,icen] for proj in subsis.listaProjTermica) \
                    + sum(modelo.capRenovCont[proj.nomeUsina,iper]*proj.seriesEolicas[icond][iper % 12]*subs.fatorPatEOL[ipat][iper%12] for proj in subsis.listaProjEOL) \
                    + sum(modelo.capRenovCont[proj.nomeUsina,iper]*proj.fatorCapacidade[iper % 12]*subs.fatorPatUFV[ipat][iper%12] for proj in subsis.listaProjUFV) \
                    + sum(modelo.capRenovCont[proj.nomeUsina,iper]*proj.fatorCapacidade[iper % 12]*subs.fatorPatBIO[ipat][iper%12] for proj in subsis.listaProjBIO) \
                    + sum(modelo.capRenovCont[proj.nomeUsina,iper]*proj.fatorCapacidade[iper % 12]*subs.fatorPatPCH[ipat][iper%12] for proj in subsis.listaProjPCH) \
                    + sum(modelo.capRenovCont[proj.nomeUsina,iper]*proj.fatorCapacidade[iper % 12]*subs.fatorPatEOF[ipat][iper%12] for proj in subsis.listaProjEOF) \
					+ sum(modelo.prodReversivel[proj.nomeUsina, ipat, iper, icen] for proj in subsis.listaProjReversivel) \
                    - sum(modelo.bombReversivel[proj.nomeUsina, ipat, iper, icen] for proj in subsis.listaProjReversivel) \
                    + modelo.deficit[isis,ipat,iper,icen]) \
                    >= modelo.demanda[isis,ipat,iper,icen] + sum(modelo.dE[iano] + modelo.dD[iano] for iano in range(0, int(iper/12)))*modelo.fatDem[isis];

        # Restricao de atendimento a demanda propria (sem intercambios)
        def restAtendeDemandaMinimaNoSistema (modelo, isis, ipat, iper, icen):
            subsis = modelo.sin.subsistemas[isis];
            subs = self.sin.subsistemas[isis];
            if (self.sin.tipoCombHidroEol == "completa"):
                icond = icen;
            elif (self.sin.tipoCombHidroEol == "intercalada"):
                icond = icen % self.sin.numEol;
            else:
                print("opcao de combinacao de series hidrologicas com eolicas nao marcada");
            # atendimento a demanda de energia
            if (iper >= self.perValidadeTemp):
                return (modelo.prodHidroExist[isis,ipat,iper,icen] \
                        + sum(modelo.prodTerm[term.nomeUsina,ipat,iper,icen] for term in subsis.listaTermica) \
                        + modelo.enPCHEx[isis,iper,icen] * subs.fatorPatPCH[ipat][iper%12] \
                        + modelo.enEOLEx[isis,iper,icen] * subs.fatorPatEOLEx[ipat][iper%12] \
                        + modelo.enUFVEx[isis,iper,icen] * subs.fatorPatUFVEx[ipat][iper%12] \
                        + modelo.enBIOEx[isis,iper,icen] * subs.fatorPatBIO[ipat][iper%12] \
                        + sum(modelo.prodHidroNova[proj.nomeUsina,ipat,iper,icen] for proj in subsis.listaProjUHE) \
                        + sum(modelo.prodTermCont[proj.nomeUsina,ipat,iper,icen] for proj in subsis.listaProjTermica) \
                        + sum(modelo.capRenovCont[proj.nomeUsina,iper]*proj.seriesEolicas[icond][iper % 12]*subs.fatorPatEOL[ipat][iper%12] for proj in subsis.listaProjEOL) \
                        + sum(modelo.capRenovCont[proj.nomeUsina,iper]*proj.fatorCapacidade[iper % 12]*subs.fatorPatUFV[ipat][iper%12] for proj in subsis.listaProjUFV) \
                        + sum(modelo.capRenovCont[proj.nomeUsina,iper]*proj.fatorCapacidade[iper % 12]*subs.fatorPatBIO[ipat][iper%12] for proj in subsis.listaProjBIO) \
                        + sum(modelo.capRenovCont[proj.nomeUsina,iper]*proj.fatorCapacidade[iper % 12]*subs.fatorPatPCH[ipat][iper%12] for proj in subsis.listaProjPCH) \
                        + sum(modelo.capRenovCont[proj.nomeUsina,iper]*proj.fatorCapacidade[iper % 12]*subs.fatorPatEOF[ipat][iper%12] for proj in subsis.listaProjEOF) \
                        + sum(modelo.prodReversivel[proj.nomeUsina, ipat, iper, icen] for proj in subsis.listaProjReversivel) \
                        - sum(modelo.bombReversivel[proj.nomeUsina, ipat, iper, icen] for proj in subsis.listaProjReversivel) \
                        + modelo.deficit[isis,ipat,iper,icen]) \
                        >= self.fatorValidadeTemp[isis] * modelo.demanda[isis,ipat,iper,icen] + sum(modelo.dE[iano] + modelo.dD[iano] for iano in range(0, int(iper/12)))*modelo.fatDem[isis];
            else:
                return Constraint.Feasible;

        # Restricao de atendimento a capacidade de potencia mais reserva operativa
        def restAtendePot (modelo, isis, iper, icen):
            subsis = modelo.sin.subsistemas[isis];
            subs = self.sin.subsistemas[isis];
            if (self.sin.tipoCombHidroEol == "completa"):
                icond = icen;
            elif (self.sin.tipoCombHidroEol == "intercalada"):
                icond = icen % self.sin.numEol;
            else:
                print("opcao de combinacao de series hidrologicas com eolicas nao marcada");
            # atendimento a demanda de ponta
            return (modelo.prodHidroExist[isis,0,iper,icen] \
                    + sum(modelo.prodReversivel[proj.nomeUsina, 0, iper, icen] for proj in subsis.listaProjReversivel) \
                    + sum(term.potUsina[iper] for term in subsis.listaTermica) \
                    + modelo.enPCHEx[isis,iper,icen] * subs.fatorPatPCH[0][iper%12] \
                    + modelo.enEOLEx[isis,iper,icen] * subs.fatorPatEOLEx[0][iper%12] \
                    + modelo.enUFVEx[isis,iper,icen] * subs.fatorPatUFVEx[0][iper%12] \
                    + modelo.enBIOEx[isis,iper,icen] * subs.fatorPatBIO[0][iper%12] \
                    + sum((1-subsis.perdasInterc[jsis][0][iper])*modelo.intercPot[jsis,isis,iper,icen] for jsis in modelo.subsistemas) - sum(modelo.intercPot[isis,jsis,iper,icen] for jsis in modelo.subsistemas) \
                    + sum(modelo.prodHidroNova[proj.nomeUsina,0,iper,icen] for proj in subsis.listaProjUHE) \
                    + sum(modelo.capTermCont[proj.nomeUsina,iper]*proj.potUsina for proj in subsis.listaProjTermica) \
                    + sum(modelo.capRenovCont[proj.nomeUsina,iper]*proj.seriesEolicas[icond][iper % 12]*subs.fatorPatEOL[0][iper%12] for proj in subsis.listaProjEOL) \
                    + sum(modelo.capRenovCont[proj.nomeUsina,iper]*proj.fatorCapacidade[iper % 12]*subs.fatorPatUFV[0][iper%12] for proj in subsis.listaProjUFV) \
                    + sum(modelo.capRenovCont[proj.nomeUsina,iper]*proj.fatorCapacidade[iper % 12]*subs.fatorPatBIO[0][iper%12] for proj in subsis.listaProjBIO) \
                    + sum(modelo.capRenovCont[proj.nomeUsina,iper]*proj.fatorCapacidade[iper % 12]*subs.fatorPatPCH[0][iper%12] for proj in subsis.listaProjPCH) \
                    + sum(modelo.capRenovCont[proj.nomeUsina,iper]*proj.fatorCapacidade[iper % 12]*subs.fatorPatEOF[0][iper%12] for proj in subsis.listaProjEOF) \
                    + modelo.deficitPot[isis,iper,icen])\
                    >= (1 + self.sin.restPot)*modelo.demanda[isis,0,iper,icen] \
                    + sum(modelo.dP[iano] + (1 + self.sin.restPot)*modelo.dD[iano] for iano in range(0, int(iper/12)))*modelo.fatDem[isis];

        # o fator de demanda dos subsistemas eh o fator carga passado como parametro da aba inicial
        modelo.fatDem = self.fatorCarga;

        # cria as restricoes 
        modelo.atendeDeman = Constraint(modelo.subsistemas, modelo.patamares, modelo.periodosTotal, modelo.condicoes, rule=restAtendeDemanda);
        modelo.atendeDemanMinimaSistema = Constraint(modelo.subsistemas, modelo.patamares, modelo.periodosTotal, modelo.condicoes, rule=restAtendeDemandaMinimaNoSistema);
        
        # so executa essas restricoes caso a opcao 'habilitar restricao de potencia' esteja marcada na aba inicial da planilha
        if self.isRestPotHabilitada:
            modelo.atendePot = Constraint(modelo.subsistemas, modelo.periodosTotal, modelo.condicoes, rule=restAtendePot);

        return;
    
    def criaRestricoesFisicas(self):
        modelo = self.modelo;
        
        # restricao de nos ficticios o somatorio dos fluxos eh nulo
        def restNoFic(modelo, isis, ipat, iper, icen):
            subsis = modelo.sin.subsistemas[isis];
            return (sum(modelo.interc[jsis,isis,ipat,iper,icen]*(1-subsis.perdasInterc[jsis][ipat][iper]) for jsis in modelo.subsistemas) \
                    - sum(modelo.interc[isis,jsis,ipat,iper,icen] for jsis in modelo.subsistemas) == 0);
        modelo.NoFic = Constraint(self.subsFic, modelo.patamares, modelo.periodosTotal, modelo.condicoes, rule=restNoFic);

        # restricao de nos ficticios o somatorio dos fluxos eh nulo
        def restNoFicPot(modelo, isis, iper, icen):
            subsis = modelo.sin.subsistemas[isis];
            return (sum(modelo.intercPot[jsis,isis, iper, icen]*(1-subsis.perdasInterc[jsis][0][iper]) for jsis in modelo.subsistemas) \
                    - sum(modelo.intercPot[isis,jsis, iper, icen] for jsis in modelo.subsistemas) == 0);
        modelo.NoFicPot = Constraint(self.subsFic, modelo.periodosTotal, modelo.condicoes, rule=restNoFicPot);

        # restricao de proibicao de deficit de energia em submercados sem carga
        def restDefEnergSub(modelo, isis, iper, icen, ipat):
            subsis = modelo.sin.subsistemas[isis];
            if modelo.demanda[isis,ipat,iper,icen] == 0:
                return (modelo.deficit[isis,ipat,iper,icen] == 0);
            else: return Constraint.Feasible;
        modelo.DefEnergSub = Constraint(modelo.subsistemas, modelo.periodosTotal, modelo.condicoes, modelo.patamares, rule=restDefEnergSub);

        # restricao de proibicao de deficit de potencia em submercados sem carga
        def restDefPotSub(modelo, isis, iper, icen):
            subsis = modelo.sin.subsistemas[isis];
            if modelo.demanda[isis,0,iper,icen] == 0:
                return (modelo.deficitPot[isis,iper,icen] == 0);
            else: return Constraint.Feasible;
        modelo.DefPotSub = Constraint(modelo.subsistemas, modelo.periodosTotal, modelo.condicoes, rule=restDefPotSub);

        return;

    def criaRestricoesCapacidades(self):
        modelo = self.modelo;
        modelo.sin = self.sin;

        # restricao de geracao maxima das termicas
        def restProdTermica (modelo, term, ipat, iper, icen):
            return (modelo.prodTerm[term,ipat,iper,icen] <= modelo.sin.listaGeralTerm[term].potUsina[iper]);
        modelo.prodTermica = Constraint(modelo.termExist, modelo.patamares, modelo.periodosTotal, modelo.condicoes, rule=restProdTermica);

        # restricao de geracao maxima das termicas continuas da expansao
        def restProdTermicaCont (modelo, term, ipat, iper, icen):
            proj = modelo.sin.listaGeralProjTerm[term]
            return (modelo.prodTermCont[term,ipat,iper,icen] <= proj.potUsina*modelo.capTermCont[term,iper]);
        modelo.prodTermicaContRest = Constraint(modelo.projTermCont, modelo.patamares, modelo.periodosTotal, modelo.condicoes, rule=restProdTermicaCont);
        
        # restricao de inflexibilidade das termicas existentes - geracao termica minima
        def restInflexibilidadeTerm(modelo, term, ipat, iper, icen):
            if (iper >= modelo.sin.listaGeralTerm[term].dataMinima):                
                return (modelo.prodTerm[term,ipat,iper,icen] >= modelo.sin.listaGeralTerm[term].potUsina[iper]*modelo.inflexTermica[term,iper]);
            else:
                return (modelo.prodTerm[term,ipat,iper,icen] == 0);
        modelo.minimaGeracaoTerm = Constraint(modelo.termExist, modelo.patamares, modelo.periodosTotal, modelo.condicoes, rule=restInflexibilidadeTerm);
    
        # restricao de inflexibilidade das termicas continuas da expansao
        def restInflexContExp (modelo, term, ipat, iper, icen):
            proj = modelo.sin.listaGeralProjTerm[term]
            return (modelo.prodTermCont[term,ipat,iper,icen] >= proj.potUsina*modelo.capTermCont[term,iper]*proj.inflexContinua[iper%12]);
        modelo.inflexExp = Constraint(modelo.projTermCont, modelo.patamares, modelo.periodosTotal, modelo.condicoes, rule=restInflexContExp);
        
        # restricao de capacidade de interligacao
        def restCapInterlig (modelo, isis, jsis, ipat, iper, icen):
            return (modelo.interc[isis,jsis,ipat,iper,icen] <= modelo.sin.subsistemas[isis].capExistente[jsis][ipat][iper] + modelo.capExpInter[isis,jsis,iper]);
        modelo.capInterlig = Constraint(modelo.subsistemas, modelo.subsistemas, modelo.patamares, modelo.periodosTotal, modelo.condicoes, rule=restCapInterlig);
        
        # restricao da capacidade de Potencia de interligacao
        def restCapInterligPot (modelo, isis, jsis, iper, icen):
            if (self.isIntercambLimitado):
                return (modelo.intercPot[isis,jsis,iper,icen] <= (modelo.sin.subsistemas[isis].capExistente[jsis][0][iper] + modelo.capExpInter[isis,jsis,iper])*modelo.sin.subsistemas[isis].limiteInterc[jsis]);
            else:
                return (modelo.intercPot[isis,jsis,iper,icen] <= modelo.sin.subsistemas[isis].capExistente[jsis][0][iper] + modelo.capExpInter[isis,jsis,iper]);
        modelo.capInterligPot = Constraint(modelo.subsistemas, modelo.subsistemas, modelo.periodosTotal, modelo.condicoes, rule=restCapInterligPot);
        
        # restricao de capacidade do agrupamento de interligacoes
        def restAgrint (modelo, agrint_ind, ipat, iper, icen):
            agrint = modelo.sin.agrints[agrint_ind];
            return (sum(modelo.interc[isis,jsis,ipat,iper,icen] for (isis,jsis) in agrint.fluxos) <= agrint.limites[ipat][iper] + sum(modelo.capExpInter[isis,jsis,iper] for (isis,jsis) in agrint.fluxos));
        modelo.agrint = Constraint(modelo.conjAgrint, modelo.patamares, modelo.periodosTotal, modelo.condicoes, rule=restAgrint);

        # restricao de capacidade de potencia do agrupamento de interligacoes
        def restAgrintPot (modelo, agrint_ind, iper, icen):
            agrint = modelo.sin.agrints[agrint_ind];
            return (sum(modelo.intercPot[isis,jsis,iper,icen] for (isis,jsis) in agrint.fluxos) <= agrint.limites[0][iper] + sum(modelo.capExpInter[isis,jsis,iper] for (isis,jsis) in agrint.fluxos));
        modelo.agrintPot = Constraint(modelo.conjAgrint, modelo.periodosTotal, modelo.condicoes, rule=restAgrintPot);

        # restricao de geracao RD - carga leve
        def restProdRDLeve (modelo, term, iper, icen):
            proj = modelo.sin.listaGeralProjTerm[term]
            if proj.nomeUsina.startswith("RD"):
                return (modelo.prodTermCont[term,modelo.sin.nPatamares-1,iper,icen] == 0);
            else:
                return Constraint.Feasible
        modelo.prodRDLeveRest = Constraint(modelo.projTermCont, modelo.periodosTotal, modelo.condicoes, rule=restProdRDLeve);

        # restricao de geracao RD - carga media
        def restProdRDMedia (modelo, term, iper, icen):
            proj = modelo.sin.listaGeralProjTerm[term]
            if proj.nomeUsina.startswith("RD"):
                return (modelo.prodTermCont[term,modelo.sin.nPatamares-2,iper,icen] == 0);
            else:
                return Constraint.Feasible
        modelo.prodRDMediaRest = Constraint(modelo.projTermCont, modelo.periodosTotal, modelo.condicoes, rule=restProdRDMedia);

        return;
    
    def criaRestricoesInvestimentos(self):
        modelo = self.modelo;
        modelo.sin = self.sin;
        
        # restricoes de data minima hidro nova       
        def restDataMinInvHidro (modelo,hidro,iper):
            # verifica a data minima de entrada
            if iper < modelo.sin.listaGeralProjUHE[hidro].dataMinima:
                return (modelo.investHidro[hidro,iper] == 0);
            else:
                return(modelo.investHidro[hidro,iper] >= 0);  
        # adiciona a restricao
        modelo.dataMinInvHidro = Constraint(modelo.projUHENova, modelo.periodosTotal, rule=restDataMinInvHidro);

        # restricao de investimento hidro unitario
        def restUmInvHidro (modelo,hidro):
            return (sum(modelo.investHidro[hidro,iper] for iper in modelo.periodosTotal) <= 1); 
        modelo.umInvHidro = Constraint(modelo.projUHENova, rule=restUmInvHidro);
        
        # restricao de investimento crescente termico, ate a data de saída se houver
        def restNaoInvTerm (modelo,term,iper):
            if (iper < modelo.sin.listaGeralProjTerm[term].dataMinima) or (iper >= modelo.sin.listaGeralProjTerm[term].dataSaida): 
                return (modelo.capTermCont[term,iper] == 0);
            else:
                return (modelo.capTermCont[term,iper] >= modelo.capTermCont[term,iper-1]);
        modelo.naoInvTerm = Constraint(modelo.projTermCont, modelo.periodosTotal, rule=restNaoInvTerm);

        # restricao de investimento crescente renovavel
        def restNaoInvRenov (modelo,renov,iper):
            if iper < modelo.sin.listaGeralProjRenov[renov].dataMinima: 
                return (modelo.capRenovCont[renov,iper] == 0); 
            else:
                return (modelo.capRenovCont[renov,iper] >= modelo.capRenovCont[renov,iper-1]); 
        modelo.naoInvRenov = Constraint(modelo.projRenovCont, modelo.periodosTotal, rule=restNaoInvRenov);

        # capacidade crescente em intercambio
        def restNaoInvInter (modelo,isis,jsis,iper):
            if iper < self.sin.perMinExpT:
                return (modelo.capExpInter[isis,jsis,iper] == 0);
            else:
                return (modelo.capExpInter[isis,jsis,iper] >= modelo.capExpInter[isis,jsis,iper-1]);
        modelo.naoInvInter = Constraint(modelo.subsistemas, modelo.subsistemas, modelo.periodosTotal, rule=restNaoInvInter);    
        
        # vinculacao entre a bidercionalidade da expansao dos fluxos
        def restInterBiDirecional(m,isis,jsis,iper):
            return (modelo.capExpInter[isis,jsis,iper] == modelo.capExpInter[jsis,isis,iper]);
        
        modelo.interBiDirecional = Constraint(modelo.subsistemas, modelo.subsistemas, modelo.periodosTotal, rule=restInterBiDirecional);    
        
        return;

    def criaRestricoesDuais(self):
        # monta as restricoes dos custos marginais de energia, potencia e duplo
        
        # CME de Energia
        def restDualE(modelo, iano):
            if iano < 3:
                return modelo.dE[iano] >= 0.01;
            else:
                return modelo.dE[iano] >= modelo.rhsDE;
        self.modelo.DualE = Constraint(self.modelo.anos, rule=restDualE);

        # CME de Potencia
        def restDualP(modelo, iano):
            return modelo.dP[iano] >=modelo.rhsDP;
        self.modelo.DualP = Constraint(self.modelo.anos, rule=restDualP);

        # CME de Duplo (Energia e Potencia)
        def restDualD(modelo, iano):
            return modelo.dD[iano] >= modelo.rhsDD;
        self.modelo.DualD = Constraint(self.modelo.anos, rule=restDualD);
        
        return;

    def criaRestricoesReducaoProblema(self):
        #restricoes para reduzir tamanho do problema
        modelo = self.modelo;
        modelo.sin = self.sin;

        # investimento em hidros apenas em janeiro/julho
        def resInvHidroJan(modelo, iproj, iper):
            if iper%12 != 0 and iper%12 != 6:
                return modelo.investHidro[iproj, iper] == 0;
            else: return Constraint.Feasible;
        modelo.invHidroJan = Constraint(modelo.projUHENova, modelo.periodosTotal, rule=resInvHidroJan);

        # investimento em termicas continuas apenas em janeiro/julho
        def resInvTermContJan(modelo, iproj, iper):
            if iper%12 != 0 and iper%12 != 6:
                return modelo.capTermCont[iproj, iper] == modelo.capTermCont[iproj, iper-1];
            else: return Constraint.Feasible;
        modelo.invTermContJan = Constraint(modelo.projTermCont, modelo.periodosTotal, rule=resInvTermContJan);

        # investimento em renovaveis continuas apenas em janeiro/julho
        def resInvRenovContJan(modelo, iproj, iper):
            if iper%12 != 0 and iper%12 != 6:
                return modelo.capRenovCont[iproj, iper] == modelo.capRenovCont[iproj, iper-1];
            else: return Constraint.Feasible;
        modelo.invRenovContJan = Constraint(modelo.projRenovCont, modelo.periodosTotal, rule=resInvRenovContJan);

        # investimento em expansao de intercambio apenas em janeiro/julho
        def resInvExpInterJan(modelo, isis, jsis, iper):
            if iper%12 != 0 and iper%12 != 6:
                return modelo.capExpInter[isis, jsis, iper] == modelo.capExpInter[isis, jsis, iper-1];
            else: return Constraint.Feasible;
        modelo.invExpInterJan = Constraint(modelo.subsistemas, modelo.subsistemas, modelo.periodosTotal, rule=resInvExpInterJan);

        # investimento em termicas continuas apenas em janeiro/julho
        def resInvReversivelJan(modelo, iproj, iper):
            if iper%12 != 0 and iper%12 != 6:
                return modelo.capReversivel[iproj, iper] == modelo.capReversivel[iproj, iper-1];
            else: return Constraint.Feasible;
        modelo.invReversivelJan = Constraint(modelo.projReversivel, modelo.periodosTotal, rule=resInvReversivelJan);

        return;

    def createIncrementais(self):
        m = self.modelo;
        
        m.capIncTermCont = Var(m.projTermCont, m.periodosTotal, domain=Reals);
        m.capIncRenovCont = Var(m.projRenovCont, m.periodosTotal, domain=NonNegativeReals);
        m.capIncReversivel = Var(m.projReversivel, m.periodosTotal, domain=NonNegativeReals);
        m.capIncExpInter = Var(m.subsistemas, m.subsistemas, m.periodosTotal, domain=NonNegativeReals);
        
        # funcao para restricoes incrementais
        def restIncrementalTerm(m,iobj,iper):
            if iper > 0:
                return m.capIncTermCont[iobj,iper] == m.capTermCont[iobj,iper] - m.capTermCont[iobj,iper-1];
            else:
                return m.capIncTermCont[iobj,iper] == m.capTermCont[iobj,iper];
        m.IncrementalTerm = Constraint(m.projTermCont, m.periodosTotal, rule=restIncrementalTerm);

        # Decisao Renovavel Continua Incremental
        def restIncrementalRenovCont(m,iobj,iper):
            if iper > 0:
                return m.capIncRenovCont[iobj,iper]  == m.capRenovCont[iobj,iper] - m.capRenovCont[iobj,iper-1];
            else:
                return m.capIncRenovCont[iobj,iper]  == m.capRenovCont[iobj,iper];
        m.IncrementalRenovCont = Constraint(m.projRenovCont, m.periodosTotal, rule=restIncrementalRenovCont);

        # Decisao Reversivel Incremental
        def restIncrementalReversivel(m,iobj,iper):
            if iper > 0:
                return m.capIncReversivel[iobj,iper]  == m.capReversivel[iobj,iper] - m.capReversivel[iobj,iper-1];
            else:
                return m.capIncReversivel[iobj,iper]  == m.capReversivel[iobj,iper];
        m.IncrementalReversivel = Constraint(m.projReversivel, m.periodosTotal, rule=restIncrementalReversivel);

        # Decisao Intercambio Incremental
        def restIncrementalExpInter(m,isis,jsis,iper):
            if iper > 0:
                return m.capIncExpInter[isis,jsis,iper]  == m.capExpInter[isis,jsis,iper] - m.capExpInter[isis,jsis,iper-1];
            else:
                return m.capIncExpInter[isis,jsis,iper]  == m.capExpInter[isis,jsis,iper];
        m.IncrementalExpInter = Constraint(m.subsistemas, m.subsistemas, m.periodosTotal, rule=restIncrementalExpInter);
        
        return;
    
    # Funcao objetivo
    def objetivo(self):
        modelo = self.modelo;
        sin = self.sin;
        modelo.sin=sin;
        tx = modelo.sin.taxaDesc;
        
        def FO (m):
            if self.isPerpetHabilitada:
                comb_perpetuo = [0 for _ in range(12)]; # inicializa com zero
                for mes in range(12):
                    iper = int(12 * (m.sin.numAnos-1) + mes) # calcula o periodo do ultimo ano
                    comb_perpetuo[mes] = (1/(1+m.sin.taxaDesc))**(iper+1) * 1/tx * 1/12 * (
                    sum(m.sin.probHidro[icen]*m.prodTerm[term, ipat, iper, icen]*m.cvuTermExist[term, iper]*m.sin.duracaoPatamar[ipat][iper] * m.sin.horasMes for term in m.termExist for ipat in m.patamares  for icen in m.condicoes) + 
                    sum(m.sin.probHidro[icen]*m.prodTermCont[term, ipat, iper, icen]*m.cvuProjTerm[term, iper]*m.sin.duracaoPatamar[ipat][iper] * m.sin.horasMes for term in m.projTermCont for ipat in m.patamares for icen in m.condicoes) +
                    sum(m.sin.probHidro[icen]*m.deficit[isis, ipat, iper , icen]*m.custoDefc[isis, ipat]*m.sin.duracaoPatamar[ipat][iper] * m.sin.horasMes for isis in m.subsistemas for ipat in m.patamares for icen in m.condicoes) +
                    sum(m.sin.probHidro[icen]*m.bombReversivel[rev, ipat, iper , icen]*sin.pldMin*m.sin.duracaoPatamar[ipat][iper] * m.sin.horasMes for rev in m.projReversivel for ipat in m.patamares for icen in m.condicoes) +
                    # termo referente ao custo de deficit
                    sum(m.sin.probHidro[icen]*m.deficitPot[isis, iper, icen]*sin.custoDefPot for isis in m.subsistemas for icen in m.condicoes)) # o custo de deficit de potencia esta em R$/MW
                    # calcula a perpetuidade do combustivel do final do horizonte - aplicando trazendo para valor presente
                return sum( (1/(1+tx))**(iper) * (
                            ( m.sin.horasMes*(  #termos referentes a operacao
                                sum(m.sin.probHidro[icen]*m.prodTerm[term, ipat, iper, icen]*m.cvuTermExist[term, iper]*m.sin.duracaoPatamar[ipat][iper] for term in m.termExist for ipat in m.patamares  for icen in m.condicoes) + 
                                sum(m.sin.probHidro[icen]*m.prodTermCont[term, ipat, iper, icen]*m.cvuProjTerm[term, iper]*m.sin.duracaoPatamar[ipat][iper] for term in m.projTermCont for ipat in m.patamares for icen in m.condicoes) +
                                sum(m.sin.probHidro[icen]*m.deficit[isis, ipat, iper , icen]*m.custoDefc[isis, ipat]*m.sin.duracaoPatamar[ipat][iper] for isis in m.subsistemas for ipat in m.patamares for icen in m.condicoes) +
                                sum(m.sin.probHidro[icen]*m.bombReversivel[rev, ipat, iper , icen]*sin.pldMin*m.sin.duracaoPatamar[ipat][iper] for rev in m.projReversivel for ipat in m.patamares for icen in m.condicoes)) +
                                sum(m.sin.probHidro[icen]*0.0005*m.interc[isis,jsis,ipat,iper,icen] for isis in m.subsistemas for jsis in m.subsistemas for ipat in m.patamares for icen in m.condicoes) +
                                sum(m.sin.probHidro[icen]*0.0005*m.intercPot[isis,jsis,iper,icen] for isis in m.subsistemas for jsis in m.subsistemas for icen in m.condicoes) +
                                sum(m.sin.probHidro[icen]*m.penalidadeGHMinExist[isis,ipat,iper,icen] for isis in m.subsistemas for ipat in m.patamares for icen in m.condicoes)*9999 +
                                sum(m.sin.probHidro[icen]*m.penalidadeGHMinNova[iuhe,ipat,iper,icen] for iuhe in m.projUHENova for ipat in m.patamares for icen in m.condicoes)*9999 +
                                # termo referente ao custo de deficit
                                sum(m.sin.probHidro[icen]*m.deficitPot[isis, iper, icen]*sin.custoDefPot for isis in m.subsistemas for icen in m.condicoes) # o custo de deficit de potencia esta em R$/MW
                            ) + \
                            ( # termos referentes a investimento - independem do cenario
                                sum(m.custoInvHidro[hidro]/tx * m.investHidro[hidro,iper] for hidro in m.projUHENova ) + 
                                sum(m.custoInvRenovCont[renov]/tx * m.capIncRenovCont[renov,iper] for renov in m.projRenovCont ) +
                                sum(1000*m.sin.subsistemas[isis].custoExpansao[jsis]/tx * m.capIncExpInter[isis,jsis,iper] for isis in m.subsistemas for jsis in range(isis, m.sin.nsis)) + 
                                sum(m.custoInvProjTerm[term]/tx * m.capIncTermCont[term,iper] / m.fdispProjTerm[term] for term in m.projTermCont ) +
                                sum(m.custoInvReversivel[iproj]/tx * m.capIncReversivel[iproj,iper] for iproj in m.projReversivel )
                            )
                        ) for iper in m.periodosTotal) + sum(comb_perpetuo[mes] for mes in range(12));
            else:
                return sum( (1/(1+m.sin.taxaDesc))**(iper) * (
                            ( m.sin.horasMes*(  #termos referentes a operacao
                                sum(m.sin.probHidro[icen]*m.prodTerm[term, ipat, iper, icen]*m.cvuTermExist[term, iper]*m.sin.duracaoPatamar[ipat][iper] for term in m.termExist for ipat in m.patamares  for icen in m.condicoes) + 
                                sum(m.sin.probHidro[icen]*m.prodTermCont[term, ipat, iper, icen]*m.cvuProjTerm[term, iper]*m.sin.duracaoPatamar[ipat][iper] for term in m.projTermCont for ipat in m.patamares for icen in m.condicoes) +
                                sum(m.sin.probHidro[icen]*m.deficit[isis, ipat, iper , icen]*m.custoDefc[isis, ipat]*m.sin.duracaoPatamar[ipat][iper] for isis in m.subsistemas for ipat in m.patamares for icen in m.condicoes) +
                                sum(m.sin.probHidro[icen]*m.bombReversivel[rev, ipat, iper , icen]*sin.pldMin*m.sin.duracaoPatamar[ipat][iper] for rev in m.projReversivel for ipat in m.patamares for icen in m.condicoes)) +
                                sum(m.sin.probHidro[icen]*0.0005*m.interc[isis,jsis,ipat,iper,icen] for isis in m.subsistemas for jsis in m.subsistemas for ipat in m.patamares for icen in m.condicoes) +
                                sum(m.sin.probHidro[icen]*0.0005*m.intercPot[isis,jsis,iper,icen] for isis in m.subsistemas for jsis in m.subsistemas for icen in m.condicoes) +
                                sum(m.sin.probHidro[icen]*m.penalidadeGHMinExist[isis,ipat,iper,icen] for isis in m.subsistemas for ipat in m.patamares for icen in m.condicoes)*9999 +
                                sum(m.sin.probHidro[icen]*m.penalidadeGHMinNova[iuhe,ipat,iper,icen] for iuhe in m.projUHENova for ipat in m.patamares for icen in m.condicoes)*9999 +
                                # termo referente ao custo de deficit
                                sum(m.sin.probHidro[icen]*m.deficitPot[isis, iper, icen]*sin.custoDefPot for isis in m.subsistemas for icen in m.condicoes) # o custo de deficit de potencia esta em R$/MW
                            ) + \
                            ( # termos referentes a investimento - independem do cenario
                                sum(m.custoInvHidro[hidro] * sum(m.investHidro[hidro,tau] for tau in range(iper+1)) for hidro in m.projUHENova ) + 
                                sum(m.custoInvRenovCont[renov] * m.capRenovCont[renov,iper] for renov in m.projRenovCont ) +
                                sum(1000*m.sin.subsistemas[isis].custoExpansao[jsis] * m.capExpInter[isis,jsis,iper] for isis in m.subsistemas for jsis in range(isis, m.sin.nsis)) + 
                                sum(m.custoInvProjTerm[term] * m.capTermCont[term,iper] / m.fdispProjTerm[term] for term in m.projTermCont ) +
                                sum(m.custoInvReversivel[iproj] * m.capReversivel[iproj,iper] for iproj in m.projReversivel )
                            )
                        ) for iper in m.periodosTotal);

        modelo.obj = Objective(rule=FO, sense=minimize);
        
        return;

    def relaxar (self):
        modelo = self.modelo;
        
        # relaxa as varaveis de decisao de investimento hidro e habilita os duais
        modelo.investHidro.domain = NonNegativeReals;
        modelo.capTermCont.domain = NonNegativeReals;
        modelo.dual = Suffix(direction=Suffix.IMPORT)
        
        return;
        
    
    def prepararDualEnergia(self):
        modelo=self.modelo;
        
        # seta os rhs para habilitar apenas o dE
        (modelo.rhsDE, modelo.rhsDD, modelo.rhsDP) =(0.3, 0, 0);
        
        return;

    def prepararDualDuplo(self):
        modelo=self.modelo;

        # seta os rhs para habilitar dE e dP
        (modelo.rhsDE, modelo.rhsDD, modelo.rhsDP) =(0, 0.3, 0);

        return;

    def prepararDualPotencia(self):
        modelo=self.modelo;

        # seta os rhs para habilitar apenas o dP
        (modelo.rhsDE, modelo.rhsDD, modelo.rhsDP) =(0, 0, 0.3);
        
        return;
