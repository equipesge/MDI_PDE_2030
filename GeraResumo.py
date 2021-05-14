import pandas as pd
import sys;
from RecebeDados import RecebeDados;
from Sistema import Sistema;
from ResumoExecutivo import ResumoExecutivo;


def preparaDF(arquivo):
    df = pd.read_csv(arquivo, delimiter=";", encoding="ISO-8859-1");

    # adiciona columa do periodo
    df=df.assign(iper=[i for i in range(len(df))])

    return df;


#pega o caminho da pasta e o nome do xls de entrada
if len(sys.argv) > 1:
    caminho = sys.argv[1] + "\\";
    planilha = sys.argv[2];
else:
    caminho = "X:\\SGE-PR~1\\02-EST~1\\33-PRO~1\\Python\\dev\\TESTE_~1\\"
    planilha = "X:\\SGE-PR~1\\02-EST~1\\33-PRO~1\\Python\\dev\\TESTE_~1\\MODELO~1.XLS"

print("Carregando dados de Entrada do MDI ...")

# carrega os dados de entrada
recebe_dados = RecebeDados(planilha);
sin = Sistema(recebe_dados, "completa");

# pega informacoes inidicias
recebe_dados.defineAba("Inicial");
pastacod = str(recebe_dados.pegaEscalar("G7"));

print("Carregando saidas ...")

# carrega os dados de saida de expansao
df = preparaDF(caminho + "saidaExpansao.txt")
dfUHE = pd.read_csv(caminho + "saidaExpansaoBinaria.txt", delimiter=" ", encoding="ISO-8859-1", header=None, skiprows=1)
dfUHE.columns=['codNW', 'none1', 'none2', 'iper']
dfCustos = pd.read_csv(caminho + "custosPeriodos.txt", delimiter=";", encoding="ISO-8859-1", header=0)

print ("Imprimindo Resumo ...")
resumo = ResumoExecutivo(sin, df, caminho, planilha, pastacod);

resumo.imprime(dfUHE, dfCustos);
