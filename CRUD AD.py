import getpass
import os
from datetime import datetime, timedelta
from ldap3 import Server, Connection, NTLM, ALL, MODIFY_REPLACE
import holidays
import subprocess
import sys

# Lista de pacotes obrigatórios
pacotes_necessarios = [
    'ldap3', # Biblioteca para interagir com o Active Directory
    'holidays', # Biblioteca para manipulação de feriados
    'getpass', # Biblioteca para entrada de senha sem exibição
    'datetime', # Biblioteca para manipulação de datas
    'os', # Biblioteca para interações com o sistema operacional
    'subprocess', # Biblioteca para execução de subprocessos
    'sys' # Biblioteca para manipulação de parâmetros do sistema
]

# Verifica e instala pacotes ausentes
def instalar_pacotes():
    for pacote in pacotes_necessarios:
        try:
            __import__(pacote)
        except ImportError:
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", pacote])
            except subprocess.CalledProcessError as e:
                sys.exit(1)
instalar_pacotes()

# Definição de variáveis globais
# Feriados do estado de São Paulo
feriados_sp = holidays.Brazil(prov='SP')
# Armazena cada busca de usuário
usuarios_encontrados = {}


# Obtém o nome do usuário logado no sistema
def get_usuario_logado():
    return os.getlogin()

# Obtém a conexão com o Active Directory baseado no usuário logado
# Primeira função a ser executada, o retorno (conexao) é usado em todas as outras funções para interagir com o AD
def get_conexao():
    usuario = get_usuario_logado()
    dominio_netbios = os.environ.get('USERDOMAIN')
    dominio_dns = os.environ.get('USERDNSDOMAIN')
    usuario_completo = f'{dominio_netbios}\\{usuario}'
    
    # Verifica se o domínio DNS foi obtido corretamente
    if not dominio_dns:
        raise Exception("Não foi possível obter o domínio DNS.")

    servidor = Server(dominio_dns, get_info=ALL)
    
    while True:
        # Solicita a senha do usuário logado
        senha = getpass.getpass("Digite sua senha do AD: ")
        try:
            # Conecta ao AD com usuário logado e senha informada
            conexao = Connection(servidor, user=usuario_completo, password=senha, authentication=NTLM, auto_bind=True)
            return conexao
        except Exception as e:
            erro_str = str(e)
            # Se a senha estiver incorreta, tenta de novo
            if "invalidCredentials" in erro_str:
                print("Credenciais incorretas.")
                continue
            # Se for outro erro, exibe mensagem e encerra
            else:
                print(f"Erro: {e}")
                sys.exit(1)

# Obtém o DN do server, que é usado como base para buscas como se fosse um prefixo
def get_base_dn(conexao):
    return conexao.server.info.other['defaultNamingContext'][0]

# Busca um usuário no Active Directory e armazena os dados encontrados em um dicionário global
def buscar_usuario(conexao):
    base_dn = get_base_dn(conexao)
    login_name = input("Digite o login ou nome do usuário: ").strip()

    # Busca por login 
    filtro_login = f'(sAMAccountName={login_name})'
    conexao.search(base_dn, filtro_login, attributes=['distinguishedName', 'displayName', 'sAMAccountName', 'memberOf'])

    # Se encontrou o usuário pelo login, armazena os dados
    if conexao.entries:
        entry = conexao.entries[0]
        dados = {
            'sAMAccountName': entry.sAMAccountName.value,
            'distinguishedName': entry.distinguishedName.value,
            'displayName': entry.displayName.value if 'displayName' in entry else entry.sAMAccountName.value,
            'grupos': entry.memberOf.values if 'memberOf' in entry else []
        }
        usuarios_encontrados[dados['sAMAccountName']] = dados
        print(f"\nUsuário encontrado:")
        print(f"  Nome de login : {dados['sAMAccountName']}")
        print(f"  Nome exibido  : {dados['displayName']}")
        print(f"  DN            : {dados['distinguishedName']}")
        return

    # Busca por nome 
    palavras = login_name.split()
    filtro_nome = '(&(objectClass=user)'
    for palavra in palavras:
        filtro_nome += f'(displayName=*{palavra}*)'
    filtro_nome += ')'

    conexao.search(base_dn, filtro_nome, attributes=['distinguishedName', 'displayName', 'sAMAccountName', 'memberOf'])

    if not conexao.entries:
        print("Nenhum usuário encontrado.")
        return

    print("\nUsuários encontrados:")
    for i, entry in enumerate(conexao.entries, start=1):
        print(f"{i}. {entry.displayName.value} ({entry.sAMAccountName.value})")

    # Seleção do usuário da lista
    while True:
        escolha = input("\nDigite o número do usuário para selecionar, 'R' para refazer a busca ou 'C' para cancelar: ").strip().upper()
        if escolha == 'R':
            return buscar_usuario(conexao)
        elif escolha == 'C':
            return
        elif escolha.isdigit() and 1 <= int(escolha) <= len(conexao.entries):
            entry = conexao.entries[int(escolha) - 1]
            dados = {
                'sAMAccountName': entry.sAMAccountName.value,
                'distinguishedName': entry.distinguishedName.value,
                'displayName': entry.displayName.value if 'displayName' in entry else entry.sAMAccountName.value,
                'grupos': entry.memberOf.values if 'memberOf' in entry else []
            }
            usuarios_encontrados[dados['sAMAccountName']] = dados
            print(f"\nUsuário selecionado:")
            print(f"  Nome de login : {dados['sAMAccountName']}")
            print(f"  Nome exibido  : {dados['displayName']}")
            print(f"  DN            : {dados['distinguishedName']}")
            return
        else:
            print("Entrada inválida. Tente novamente.")

# Lista todos os usuários ativos do Active Directory
def lista_usuarios_ativos(conexao):
    base_dn = get_base_dn(conexao)
    filtro = '(&(objectClass=user)(objectCategory=person)(userAccountControl:1.2.840.113556.1.4.803:=512))'
    
    conexao.search(base_dn, filtro, attributes=['sAMAccountName', 'displayName', 'distinguishedName'])
    
    if not conexao.entries:
        print("Nenhum usuário ativo encontrado.")
        return

    print("\nUsuários ativos encontrados:")
    for entry in conexao.entries:
        print(f"{entry.displayName.value} ({entry.sAMAccountName.value}) - DN: {entry.distinguishedName.value}")

# Altera o campo 'escritório' de um usuário previamente buscado
def alterar_escritorio(conexao):
    if not usuarios_encontrados:
        print("Realize uma busca antes de alterar.")
        return

    print("\nUsuários encontrados:")
    for i, (login, dados) in enumerate(usuarios_encontrados.items(), start=1):
        print(f"{i}. {dados['displayName']} ({login})")

    try:
        escolha = int(input("Selecione o número do usuário: "))
        if escolha < 1 or escolha > len(usuarios_encontrados):
            print("Número inválido.")
            return
    except ValueError:
        print("Entrada inválida.")
        return

    login_selecionado = list(usuarios_encontrados.keys())[escolha - 1]
    usuario = usuarios_encontrados[login_selecionado]
    novo_valor = input(f"Defina o novo local de trabalho de {usuario['displayName']}: ")

    resultado = conexao.modify(usuario['distinguishedName'], {
        'physicalDeliveryOfficeName': [(MODIFY_REPLACE, [novo_valor])]
    })

    if resultado:
        print(f"Escritório alterado para: {novo_valor}")
        registrar_log_acao(conexao, usuario['distinguishedName'], "Alteração")
    else:
        print("Erro ao alterar:", conexao.result)

# Conta o número de membros em grupos específicos do Active Directory
# Alterar posteriormente para buscar todos os grupos de organização no AD
def contar_membros_grupos(conexao):
    base_dn = get_base_dn(conexao)
    
    # Lista dos nomes exatos dos grupos
    nomes_grupos = [
        "01 - Presidente",
        "02 - Diretores",
        "03 - Chefe de Gabinete",
        "04 - Ouvidor",
        "05 - Assessores",
        "06 - Superintendentes",
        "07 - Gerentes",
        "08 - Líderes",
        "09 - Demais Funcionários",
        "10 - Estagiários",
        "11 - Aprendizes",
        "12 - Prestadores",
        "13 - Gerentes Regionais",
        "14 - Comitê de Entregas"
    ]
    
    print("\nContagem de membros por grupo:")
    for nome_grupo in nomes_grupos:
        filtro = f"(&(objectClass=group)(cn={nome_grupo}))"
        conexao.search(base_dn, filtro, attributes=['member'])
        if conexao.entries:
            grupo = conexao.entries[0]
            membros = grupo.member.values if 'member' in grupo else []
            print(f"- {nome_grupo}: {len(membros)} membro(s)")
        else:
            print(f"- {nome_grupo}: grupo não encontrado.")

# Calcula o segundo dia útil a partir de uma data base, considerando feriados e fins de semana
def ajustar_dia_util(data: datetime) -> datetime:
    while True:
        dia_semana = data.weekday()  # 0=Segunda, ..., 6=Domingo
        if data in feriados_sp:
            data += timedelta(days=1)
            continue
        if dia_semana == 4:  # Sexta-feira
            data += timedelta(days=4)  # Vai pra terça
            continue
        if dia_semana == 0:  # Segunda-feira
            data += timedelta(days=1)  # Vai pra terça
            continue
        if dia_semana >= 5:  # Sábado ou domingo
            data += timedelta(days=1)
            continue
        break  # Data válida
    return data


# Renova a conta de um usuário previamente buscado, adicionando 90 ou 180 dias
# Se o usuário for membro do grupo "12 - Prestadores", adiciona 90 dias. Se não, adiciona 180 dias.
def renovar_conta(conexao):
    if not usuarios_encontrados:
        print("Nenhum usuário foi buscado ainda.")
        return

    print("\nUsuários buscados:")
    for i, (login, dados) in enumerate(usuarios_encontrados.items(), start=1):
        print(f"{i}. {dados['displayName']} ({login})")

    try:
        escolha = int(input("Selecione o número do usuário para renovar a conta: "))
        if escolha < 1 or escolha > len(usuarios_encontrados):
            print("Número inválido.")
            return
    except ValueError:
        print("Entrada inválida.")
        return

    login_selecionado = list(usuarios_encontrados.keys())[escolha - 1]
    usuario = usuarios_encontrados[login_selecionado]

    # Verifica se está no grupo "12 - Prestadores"
    membro_prestador = any("12 - Prestadores" in grupo for grupo in usuario['grupos'])
    dias_para_adicionar = 90 if membro_prestador else 180

    data_futura = datetime.today() + timedelta(days=dias_para_adicionar)
    data_ajustada = ajustar_dia_util(data_futura)

    # Converte para formato do AD (intervalo de 100 nanossegundos desde 1601-01-01)
    ticks = int((data_ajustada - datetime(1601, 1, 1)).total_seconds() * 10**7)

    resultado = conexao.modify(usuario['distinguishedName'], {
        'accountExpires': [(MODIFY_REPLACE, [str(ticks)])]
    })

    if resultado:
        print(f"Conta {usuario['sAMAccountName']} renovada até {data_ajustada.strftime('%d/%m/%Y')} ({dias_para_adicionar} dias).")
        registrar_log_acao(conexao, usuario['distinguishedName'], "Renovação")
    else:
        print("Erro ao renovar conta:", conexao.result)


def registrar_log_acao(conexao, dn, tipo_acao):
    chamado = input("Informe o número do chamado: ").strip()
    data_hoje = datetime.today().strftime('%d/%m/%Y')
    nova_linha = f"{data_hoje} - {tipo_acao} - {chamado}"

    # Lê o conteúdo atual do campo 'info'
    conexao.search(dn, '(objectClass=*)', attributes=['info'])
    atual = conexao.entries[0]['info'].value if 'info' in conexao.entries[0] else ""

    # Adiciona nova linha ao final
    novo_conteudo = ( "\r\n" + atual + "\r\n" + nova_linha).strip() if atual else nova_linha

    resultado = conexao.modify(dn, {
        'info': [(MODIFY_REPLACE, [novo_conteudo])]
    })

    if resultado:
        print(f"Log adicionado no campo 'Observação': {nova_linha}")
    else:
        print("Falha ao registrar ação:", conexao.result)



def menu():
    conexao = get_conexao()
    while True:
        print("\n=== MENU AD ===")
        print("1. Buscar usuário no AD")
        print("2. Alterar campo 'escritório' de um usuário buscado")
        print("3. Renovar conta de usuário")
        print("4. Contar membros dos grupos principais")
        print("5. Listar usuários ativos")
        print("5. Sair")

        opcao = input("Escolha uma opção: ")
        if opcao == '1':
            buscar_usuario(conexao)
        elif opcao == '2':
            alterar_escritorio(conexao)
        elif opcao == '3':
            renovar_conta(conexao)
        elif opcao == '4':
            contar_membros_grupos(conexao)
        elif opcao == '5':
            lista_usuarios_ativos(conexao)
        elif opcao == '6':
            print("Encerrando...")
            break
        else:
            print("Opção inválida.")


if __name__ == "__main__":
    menu()
