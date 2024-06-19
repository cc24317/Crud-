import pandas as pd
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
        return True #Retorna verdadeiro se a conexão for bem sucedida
    except Exception as e:
        print(f"Erro ao conectar: {str(e)}")
        return False #Retorna falso se acontecer algum erro na conexão 
    
def planilha():
    if connect():
        try:
            cursor1 = connection.cursor() 
            cursor1.execute("SELECT nomeMecanico FROM DadosDaRoca.mecanico")
            nomeMecanico = cursor1.fetchall()[0]  
            cursor1.close()

            cursor2 = connection.cursor()
            cursor2.execute("SELECT placa FROM DadosDaRoca.veiculo")
            placaVeiculo = cursor2.fetchall()[0]  
            cursor2.close()

            cursor3 = connection.cursor()
            cursor3.execute("SELECT tipoManutencao FROM DadosDaRoca.ordemServico")
            tipoOrdemServico = cursor3.fetchall()[0]  
            cursor3.close()
            
            cursor4 = connection.cursor()
            cursor4.execute("SELECT numeroOrdemServico FROM DadosDaRoca.ordemServico")
            numeroOrdemServico = cursor4.fetchall()[0]  
            cursor4.close()

            cursor5 = connection.cursor()
            cursor5.execute("SELECT tempoEstimado FROM DadosDaRoca.ordemServico")
            horarioExecucao = cursor5.fetchall()[0]  
            cursor5.close()

            informacoes = {
                'Nome do Mecânico': [nomeMecanico],
                'Número do Veículo': [placaVeiculo],
                'Tipo da Ordem de Serviço': [tipoOrdemServico],
                'Número da Ordem de Serviço': [numeroOrdemServico],
                'Horário': [horarioExecucao]
            }

            dataframe = pd.DataFrame(informacoes)
            dataframe.to_excel('Planilha.xlsx', index=False)
            print("DataFrame salvo na planilha Excel 'Planilha.xlsx'")

            # Fechando cursores e conexão
            connection.close()

        except Exception as e:
            print(f"Erro ao gerar planilha: {str(e)}")
    else:
        print("Não foi possível conectar ao banco de dados.")

if __name__ == "__main__":
    planilha()

    # Salvando o DataFrame em um arquivo CSV
    # df.to_excel('', index=False)
    # print("DataFrame salvo no arquivo CSV 'dados.csv'")

def escolhaOrdemServico():
    if connect():
        cursor = connection.cursor()

        # Exemplo de consulta para buscar a próxima ordem de serviço disponível
        cursor.execute("SELECT numeroOrdemServico FROM DadosDaRoca.ordemServico")
        resultado = cursor.fetchall()

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

    else:
        print("Não foi possível conectar ao banco de dados.")

if __name__ == "__main__":
    container()
