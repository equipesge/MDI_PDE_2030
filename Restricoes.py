from RecebeDados import RecebeDados;
import jsonpickle;

class Restricoes:

    def __init__(self, plan_dados, numAnos, SIN):
        
        # carrega a planilha
        self.recebe_dados = plan_dados;
        self.numAnos = numAnos;
        self.SIN = SIN;

        # carrega as informacoes da planilha
        self.load();
                
        return;

    def load(self):
        # inicializa os hashs
        self.Step = [];self.LimiteAno =[];self.Igualdade =[];self.IgualdadeMax =[];self.LimiteIncAno=[];self.Proporcao=[];

        # define a aba
        self.recebe_dados.defineAba("Restricoes_Adicionais");

        # percorre todas as linhas
        iRest = 1;
        while (self.recebe_dados.pegaEscalar("A1", lin_offset=iRest)is not None):
            # pega os parametros
            tipoProj = self.recebe_dados.pegaEscalar("A1", lin_offset=iRest);
            codProj = self.recebe_dados.pegaEscalar("B1", lin_offset=iRest);
            tipoRest = self.recebe_dados.pegaEscalar("C1", lin_offset=iRest);
            valores = [self.recebe_dados.pegaEscalar("D1", lin_offset=iRest, col_offset=col) for col in range(0,self.numAnos)];
            mes = self.recebe_dados.pegaEscalar("D1", lin_offset=iRest, col_offset=self.numAnos);
            if not(mes == None):
                mes = int(mes)

            # cria a restricao
            rest = Restricao(tipoProj, codProj, tipoRest, valores, self.SIN, mes);

            # adiciona o tipo de restricao
            if (tipoRest == "LimiteAno"):
                self.LimiteAno.append(rest);
            if (tipoRest == "Step"):
                self.Step.append(rest);
            if (tipoRest == "Igualdade"):
                self.Igualdade.append(rest);
            if (tipoRest == "IgualdadeMax"):
                self.IgualdadeMax.append(rest);
            if (tipoRest == "LimiteIncAno"):
                self.LimiteIncAno.append(rest);
            if (tipoRest == "Proporcao"):
                self.Proporcao.append(rest);

            iRest = iRest + 1;

        return;

class Restricao:
    def __init__(self, tipoProj, codProj, tipoRest, valores, SIN, mes):
        # pega os parametros que devem ser apenas armazenados
        self.tipoProj = tipoProj;
        self.tipoRest = tipoRest;
        if mes is not None:
            mes += -1;
        self.mes = mes;

        # pega a lista de projetos
        codProj = str(codProj);
        if tipoProj == "RenovCont":
            self.listaProj = [SIN.listaIndGeralProjRenov[int(float(ind)-1)].nomeUsina for ind in codProj.split(";")];
        
        if tipoProj == "Reversivel":
            self.listaProj = [SIN.listaIndGeralProjReversivel[int(float(ind)-1)].nomeUsina for ind in codProj.split(";")];
			
        if tipoProj == "Hidro": 
            self.listaProj = [SIN.listaIndGeralProjUHE[int(float(ind))].nomeUsina for ind in codProj.split(";")];

        if tipoProj == "Term":
            # o projeto de termica nao subtrai 1 porque o numero eh externo
            self.listaProj = [SIN.listaIndGeralProjTerm[int(float(ind))].nomeUsina for ind in codProj.split(";")];
            
        self.valores = valores;

        # monta a parte dos anos
        self.montaAnos();
        
        if (tipoProj == "Term"):
            # no caso das termicas continuas tem que deduzir o teif/ip - quando eh step e tem mais de uma pega o teif/ip da primeira            
            proj = SIN.listaGeralProjTerm[self.listaProj[0]];
            if (tipoRest == "Step"):
                s=";"
                self.valores[self.anoInicial] = s.join([str(float(v)*proj.fdisp) for v in self.valores[self.anoInicial].split(";")])
            else:
                for i in range(self.anoInicial, self.anoFinal+1):
                    self.valores[i] = float(self.valores[i])*proj.fdisp;

        # caso seja uma restricao do tipo limite complementa as informacoes necessarios
        if (tipoRest == "Step"): 
            self.montaStep();
        if (tipoRest == "Proporcao"): 
            self.montaProporcao();
            
        # caso de UHE consta o mes
        if ((tipoRest == "Igualdade") and (tipoProj == "Hidro")): 
            self.mes = int(self.valores[self.anoInicial]-1);
        if ((tipoRest == "IgualdadeMax") and (tipoProj == "Hidro")): 
            self.mes = int(self.valores[self.anoInicial]-1);

        return;

    def montaStep(self):
        # configura os parametros limites do step
        (self.val_min, self.val_max) = self.valores[self.anoInicial].split(";");
        self.val_min = float(self.val_min); self.val_max = float(self.val_max);
        return;

    def montaAnos(self):
        # o primeiro elemento que tem valor diferente de nulo
        iAno = 0;
        numAnos = len(self.valores);
        while ((iAno < numAnos) and (self.valores[iAno] is None)) : iAno = iAno +1;

        # joga para anoinicial
        self.anoInicial = iAno;

        # pega o ano final
        while ((iAno < numAnos) and (self.valores[iAno] is not None)) : iAno = iAno +1;
        self.anoFinal = iAno-1;

        return;

    def montaProporcao(self):
    
        # inicializa o dict
        self.valoresProj = [[0 for x in range(self.anoFinal-self.anoInicial+1)] for y in range(len(self.listaProj))];
        
        # neste caso os valores tem que estar num dict por periodo
        for iano in range(self.anoInicial, self.anoFinal+1):
            valores = self.valores[iano].split(";")
            for iusi in range(len(self.listaProj)):
                self.valoresProj[iusi][iano-self.anoInicial] = float(valores[iusi]);
                
        return;