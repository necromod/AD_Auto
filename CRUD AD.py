import getpass
import os
from datetime import datetime, timedelta
from ldap3 import Server, Connection, NTLM, ALL, MODIFY_REPLACE
import holidays
import socket
import subprocess
import sys

# Lista de pacotes obrigatórios
pacotes_necessarios = [
    'ldap3',
    'pycryptodome',
    'holidays',
    'getpass',
    'socket',
    'datetime',
    'os',
    'subprocess',
    'sys'
]

# Verifica e instala pacotes ausentes
def instalar_pacotes():
    for pacote in pacotes_necessarios:
        try:
            __import__(pacote)
        except ImportError:
            print(f"Instalando pacote: {pacote}")
            subprocess.check_call([sys.executable, "-m", "pip", "install", pacote])

instalar_pacotes()
# Feriados do estado de São Paulo
feriados_sp = holidays.Brazil(prov='SP')

# Armazena cada busca de usuário
usuarios_encontrados = {}

def get_usuario_logado():
    return os.getlogin()

# Obtém a conexão com o Active Directory baseado no usuário logado
def get_conexao():
    usuario = get_usuario_logado()
    dominio_netbios = os.environ.get('USERDOMAIN')
    dominio_dns = os.environ.get('USERDNSDOMAIN')
    usuario_completo = f'{dominio_netbios}\\{usuario}'
    senha = getpass.getpass("Digite sua senha do AD: ")

    if not dominio_dns:
        raise Exception("Não foi possível obter o domínio DNS.")

    servidor = Server(dominio_dns, get_info=ALL)
    conexao = Connection(servidor, user=usuario_completo, password=senha, authentication=NTLM, auto_bind=True)
    return conexao

# Obtém o DN do server, que é usado como base para buscas como se fosse um prefixo
def get_base_dn(conexao):
    return conexao.server.info.other['defaultNamingContext'][0]

# Busca um usuário no Active Directory e armazena os dados encontrados em um dicionário global
def buscar_usuario(conexao):
    base_dn = get_base_dn(conexao)
    nome_usuario = input("Digite o nome de login (sAMAccountName) do usuário a buscar: ").strip()

    filtro = f'(sAMAccountName={nome_usuario})'
    conexao.search(base_dn, filtro, attributes=['distinguishedName', 'displayName', 'sAMAccountName', 'memberOf'])

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
    else:
        print("Usuário não encontrado no AD.")

def alterar_escritorio(conexao):
    if not usuarios_encontrados:
        print("Nenhum usuário foi buscado ainda.")
        return

    print("\nUsuários buscados:")
    for i, (login, dados) in enumerate(usuarios_encontrados.items(), start=1):
        print(f"{i}. {dados['displayName']} ({login})")

    try:
        escolha = int(input("Selecione o número do usuário para alterar: "))
        if escolha < 1 or escolha > len(usuarios_encontrados):
            print("Número inválido.")
            return
    except ValueError:
        print("Entrada inválida.")
        return

    login_selecionado = list(usuarios_encontrados.keys())[escolha - 1]
    usuario = usuarios_encontrados[login_selecionado]
    novo_valor = input(f"Digite o novo valor para o campo 'escritório' de {usuario['displayName']}: ")

    resultado = conexao.modify(usuario['distinguishedName'], {
        'physicalDeliveryOfficeName': [(MODIFY_REPLACE, [novo_valor])]
    })

    if resultado:
        print(f"Campo 'escritório' alterado com sucesso para: {novo_valor}")
        registrar_log_acao(conexao, usuario['distinguishedName'], "Alteração")
    else:
        print("Erro ao alterar:", conexao.result)

def contar_membros_grupos(conexao):
    base_dn = get_base_dn(conexao)
    
    # Lista dos nomes exatos dos grupos (os mesmos do seu print)
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


def segundo_dia_util_apartir(base_data: datetime, dias_adicionais: int) -> datetime:
    data_alvo = base_data + timedelta(days=dias_adicionais)
    dias_uteis_encontrados = 0
    while dias_uteis_encontrados < 2:
        data_alvo += timedelta(days=1)
        if data_alvo.weekday() >= 5 or data_alvo in feriados_sp:
            continue
        dias_uteis_encontrados += 1
    return data_alvo

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
    nova_data = segundo_dia_util_apartir(datetime.today(), dias_para_adicionar)

    # Converte para formato do AD (intervalo de 100 nanossegundos desde 1601-01-01)
    ticks = int((nova_data - datetime(1601, 1, 1)).total_seconds() * 10**7)

    resultado = conexao.modify(usuario['distinguishedName'], {
        'accountExpires': [(MODIFY_REPLACE, [str(ticks)])]
    })

    if resultado:
        print(f"Conta de {usuario['displayName']} renovada até {nova_data.strftime('%d/%m/%Y')} ({'90' if membro_prestador else '180'} dias + 2º dia útil).")
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
        print(f"Ação registrada com sucesso no campo 'Observações': {nova_linha}")
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
            print("Encerrando...")
            break
        else:
            print("Opção inválida.")


if __name__ == "__main__":
    menu()
