import pandas
import pyodbc

def connect() -> bool: #Função para conectar o banco de dados com o programa
    global connection #variavel capaz da ser usada ou modificada em qualquer outra função do código
    try:

        connection = pyodbc.connect(
        driver = "{SQL Server}", #fabricante
        server = "143.106.250.84", #Endereço ip da maquina onde esta o banco de dados
        database = "BD24317", #banco de dados que será acessado
        uid = "BD24317", #login
        pwd = "BD24317" #senha
        )
        return True #Retorna veidadeiro se a conexão for bem sucedida
    except Exception as e:
        print(f"Erro ao conectar: {str(e)}")
        return False #Retorna falso se acontecer algum erro na conexão 
    
def planilha(): #Função para gerar planilha com os dados correspondentes ao banco de dados conectado
    if connect():
        try:
            cursor1 = connection.cursor()
            nomeMecanicoQuery = "SELECT * FROM DadosDaRoca.mecanico WHERE nome='nome'"
            cursor1.execute(nomeMecanicoQuery) #Executa a consulta no SQL definida na variável
            nomeMecanico = cursor1.fetchall() #Obtém os resultados da consulta e armazana na variável

            cursor2 = connection.cursor()
            placaVeiculoQuery = "SELECT * FROM DadosDaRoca.veiculo WHERE placa='placa'"
            cursor2.execute(placaVeiculoQuery)
            placaVeiculo = cursor2.fetchall()

            cursor3 = connection.cursor()
            ordemServicoQuery = "SELECT * FROM DadosDaRoca.ordemServico WHERE ordemServico='codigoOrdemServico'"
            cursor3.execute(ordemServicoQuery)
            ordemServico = cursor3.fetchall()

            cursor4 = connection.cursor()
            tipoManutencaoQuery = "SELECT * FROM DadosDaRoca.ordemServico WHERE manutencao='tipoManutencao'"
            cursor4.execute(tipoManutencaoQuery)
            tipoManutencao = cursor4.fetchall()

            cursor5 = connection.cursor()
            horarioExecucaoQuery = "SELECT * FROM DadosDaRoca.ordemServico WHERE horario='horarioInicio'"
            cursor5.execute(horarioExecucaoQuery)
            horarioExecucao = cursor5.fetchall()

            informacoes = {
                'Nome do Mecânico': [nomeMecanico],
                'Número do Veículo': [placaVeiculo],
                'Tipo da Ordem de Serviço': [ordemServico],
                'Número da Ordem de Serviço': [tipoManutencao],
                'Horário': [horarioExecucao]
            }

            dataframe = pandas.DataFrame(informacoes)

            # Salvando o DataFrame em um arquivo Excel
            dataframe.to_excel(r'c:\Users\u24317\Desktop\PP\darocaProjeto\Planilha.xlsx', index=False)
            print("DataFrame salvo na planilha Excel 'Planilha.xlsx'")

        except Exception as e:
                print(f"Erro ao gerar planilha: {str(e)}")
    else:
        print("Não foi possível conectar ao banco de dados.")

    # Salvando o DataFrame em um arquivo CSV
    # df.to_excel('', index=False)
    # print("DataFrame salvo no arquivo CSV 'dados.csv'")

def escolhaOrdemServico():
    if connect():
        cursor = connection.cursor()

        # Exemplo de consulta para buscar a próxima ordem de serviço disponível
        cursor.execute("SELECT numeroOrdemServico FROM DadosDaRoca.ordemServico")
        resultado = cursor.fetchone()

        if resultado:
            return resultado[0]  # Retorna o número da ordem de serviço encontrada
        else:
            print("Nenhuma há nenhuma ordem de serviço disponível no momento.")
            return None  # Ou outro valor para indicar que não há ordens disponíveis
    else:
        print("Não foi possível conectar ao banco de dados.")
        return None

# Função para executar a lógica principal do programa
def container():
    if connect():
        cursor = connection.cursor()

        # Exemplo de busca de dados do banco de dados
        cursor.execute("SELECT numeroOrdemServico FROM DadosDaRoca.ordemServico")
        ordemServico = [row[0] for row in cursor.fetchall()]

        cursor.execute("SELECT nomeMecanico FROM DadosDaRoca.mecanico")
        mecanico = [row[0] for row in cursor.fetchall()]

        centroDistOS = '1'
        centroDistMEC = '2'
        horasTrabalhadas = [0] * len(mecanico)  # Inicializa as horas trabalhadas de cada mecânico
        roteiro = [[] for _ in range(len(mecanico))]  # Planilha final
        tempoOrdemServico = 2  # Tempo estimado de cada ordem de serviço (exemplo fixo)

        while max(horasTrabalhadas) < 8:  # Enquanto nenhum mecânico ultrapassar 8 horas de trabalho
            for i, mecanico in enumerate(mecanico):
                ordemServico = escolhaOrdemServico()
                if centroDistOS != centroDistMEC:
                    continue
                if horasTrabalhadas[i] >= 8:
                    continue
                if horasTrabalhadas[i] + tempoOrdemServico[ordemServico - 1] > 8:
                    continue
                if 3 <= horasTrabalhadas[i] <= 5:  # Se o mecânico tem entre 3 e 5 horas de trabalho disponíveis
                    roteiro[i].append("Almoço")
                    horasTrabalhadas[i] += 1  # Incrementa 1 hora por conta do almoço
                roteiro[i].append(ordemServico)  # Registra a ordem de serviço no roteiro
                horasTrabalhadas[i] += tempoOrdemServico[ordemServico - 1]  # Adiciona o tempo estimado da ordem de serviço

        # Exibe o roteiro final de cada mecânico
        for i, m in enumerate(mecanico):
            print("Mecânico", m, ":", roteiro[i])

        # Gerar planilha (ainda precisa ser implementado)
        # planilha()

    else:
        print("Não foi possível conectar ao banco de dados.")

if __name__ == "__main__":
    container()
