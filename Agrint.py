from RecebeDados import RecebeDados;

class Agrint:
    
    # obs: AGRINT significa agrupamento de interligacoes
    def __init__(self, recebe_dados, iagrint, nper, nperPos, nPatamares):
        
        # define fonte_dados como o objeto da classe RecebeDados e o index interno do AgrInt
        self.fonte_dados = recebe_dados;
        self.indexAgrintInterno = iagrint;
        self.nper = nper;
        self.nperPos = nperPos;
        self.npat = nPatamares;
        
        # declara os vetores com os dados
        self.fluxos = [];
        self.limites = [[0 for iper in range(0,self.nper)] for ipat in range(0, self.npat)];
        
        # chama o metodo que importa os agrupamentos
        self.importaAgrupamento();
               
        return;
        
    def construirLista(self, nSubs):
        
        self.nsis = nSubs
        # inicializa com zeros
        lista = [[0 for isis in range(self.nsis)] for jsis in range(self.nsis)];
        
        # insere os fluxos
        for (isis,jsis) in self.fluxos:
            lista[isis][jsis] = 1;
            
        return lista;
    
    def importaAgrupamento(self):
        
        # reforca a aba da vez e cria as variaveis auxiliares
        self.fonte_dados.defineAba("AGRINT-Grupos");
        linhaOffset = 4*self.indexAgrintInterno;
        colunaOffset = 0;
        
        # importa as informacoes de fato
        self.indexAgrintExterno = self.fonte_dados.pegaEscalar("A2", lin_offset=self.indexAgrintInterno);
        self.numInterligacoes = int(self.fonte_dados.pegaEscalar("B2", lin_offset=self.indexAgrintInterno));
        for i in range(self.numInterligacoes):
            vetDePara = [int(self.fonte_dados.pegaEscalar("C2", lin_offset=self.indexAgrintInterno, col_offset = colunaOffset)-1), int(self.fonte_dados.pegaEscalar("C2", lin_offset=self.indexAgrintInterno, col_offset = colunaOffset+1)-1)];
            self.fluxos.append(vetDePara);
            colunaOffset += 2;

        # pega os limites desse agrupamento
        # muda para a aba de limites
        self.fonte_dados.defineAba("AGRINT-Limites");
        for ipat in range(0, self.npat):
            self.limites[ipat] = self.fonte_dados.pegaVetor("C2", lin_offset = linhaOffset, direcao = "horizontal", tamanho=self.nper);
            linhaOffset += 1;
            # repete o ultimo ano para o periodo pós
            for iper in range(self.nperPos):
                self.limites[ipat].append(self.limites[ipat][self.nper - 12 + iper%12]);
       
        return;