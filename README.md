# AD_Auto - Sistema de Auditoria do Active Directory

Projeto de automação do Active Directory versão Windows Server 2016, manipulação realizada em Python via LDAP3 para geração de relatórios completos e auditoria de usuários.

## 📋 Descrição

O **AD_Auto** é uma ferramenta desenvolvida em Python que permite conectar-se ao Active Directory e gerar relatórios detalhados em Excel sobre usuários, contas ativas/inativas, e realizar auditorias específicas com critérios personalizados.

## 🔧 Funcionalidades

### 📊 Relatórios Disponíveis

1. **Contas de usuário ativas** - Lista todas as contas habilitadas no domínio
2. **Contas desabilitadas a partir de 01/04/2024** - Usuários desabilitados recentemente
3. **Contas criadas em 2024** - Novos usuários criados no ano atual
4. **Contas desabilitadas em 2024** - Usuários desabilitados durante o ano
5. **Relação de e-mails** - Diretório completo com Nome, E-mail e Cargo
6. **Auditoria 2024** - Relatório completo com critérios específicos de auditoria

### 🎯 Características Principais

- **Conexão automática** ao Active Directory usando credenciais do usuário logado
- **Busca paginada** para lidar com grandes volumes de dados
- **Geração de planilhas Excel** com formatação profissional
- **Critérios de auditoria** personalizados para compliance
- **Interface interativa** com menu de opções
- **Tratamento robusto de erros** e validações
- **Análise detalhada** de dados do AD para debug

## 📋 Pré-requisitos

### Sistema Operacional
- Windows Server 2016+ ou Windows 10+
- Conectividade com Active Directory

### Python
- Python 3.6+
- Bibliotecas Python:
  - `ldap3` - Conectividade LDAP
  - `openpyxl` - Geração de planilhas Excel
  - `datetime` - Manipulação de datas
  - `os`, `getpass`, `socket`, `ssl` - Bibliotecas padrão

### Permissões
- Usuário deve ter permissões de leitura no Active Directory
- Acesso às OUs de usuários do domínio

## 🚀 Instalação

### 1. Clone o repositório
```bash
git clone https://github.com/usuario/AD_Auto.git
cd AD_Auto
```

### 2. Instale as dependências
```bash
pip install ldap3 openpyxl
```

*Nota: O sistema instalará automaticamente as bibliotecas necessárias se não estiverem disponíveis.*

## 📁 Estrutura do Projeto

```
AD_Auto/
├── List_AD.py          # Script principal com menu interativo
├── Auditoria_2024.py   # Sistema de auditoria completo
├── Busca_AD.py         # Ferramenta de debug para usuários específicos
└── README.md           # Documentação do projeto
```

## 📖 Como Usar

### 1. Execução do programa principal
```bash
python List_AD.py
```

### 2. Auditoria completa
```bash
python Auditoria_2024.py
```

### 3. Debug de usuário específico
```bash
python Busca_AD.py
```

### 4. Conexão ao Active Directory
- O sistema detectará automaticamente:
  - Usuário logado no Windows
  - Domínio NetBIOS e DNS
  - Configurações de conectividade

- Será solicitada a senha do usuário para autenticação

### 5. Seleção de relatório
- Escolha uma das 6 opções disponíveis no menu
- O relatório será gerado automaticamente em Excel
- O arquivo será aberto automaticamente após a criação

### 6. Exemplo de uso
```
🔍 MENU DE RELATÓRIOS
========================================
1️⃣  Contas de usuário ativas
2️⃣  Contas desabilitadas a partir de 01/04/2024  
3️⃣  Contas de usuários criadas somente em 2024
4️⃣  Contas de usuários desabilitadas somente em 2024
5️⃣  Relação de todos os e-mails (Nome, E-mail, Cargo)
6️⃣  Auditoria 2024 (Critério completo original)
0️⃣  Sair
========================================

🔍 Escolha uma opção (0-6): 1
```

## 🎯 Critérios de Auditoria 2024

### Contas Ativas
- **Incluídas**: Todas as contas com status ativo no Active Directory

### Contas Inativas
- **Incluídas SE**:
  - Data de criação >= 01/01/2024 **OU**
  - Último logon >= 01/01/2024 e < 01/01/2025
  - Contas sem data de expiração válida mas com logon em 2024

### Exclusões
- Contas inativas antigas sem atividade em 2024
- Contas criadas em 2025 ou posterior
- Contas com último logon em 2025 ou posterior

## 📊 Estrutura dos Relatórios

### Colunas Padrão
- **Login**: sAMAccountName do usuário
- **Nome**: displayName do usuário
- **E-mail**: endereço de email corporativo
- **Cargo**: title/função do usuário
- **Status**: Ativo/Inativo
- **Data de Criação**: whenCreated
- **Data de Expiração**: accountExpires (quando aplicável)
- **Último Logon**: lastLogon/lastLogonTimestamp

### Formatação Excel
- Cabeçalhos com formatação profissional
- Informações de geração (data/hora, critérios)
- Estatísticas do relatório
- Colunas com largura otimizada
- Cores e estilos padronizados

## 🔧 Configurações Técnicas

### Conectividade LDAP
- **Porta 389**: LDAP padrão
- **Porta 636**: LDAPS com SSL
- **Paginação**: 1000 registros por página
- **Timeout**: Configurável por operação
- **Autenticação**: NTLM com credenciais do usuário logado

### Filtros LDAP
```ldap
# Usuários ativos
(&(objectClass=user)(!(sAMAccountName=*$))(!(userAccountControl:1.2.840.113556.1.4.803:=2)))

# Usuários inativos
(&(objectClass=user)(!(sAMAccountName=*$))(userAccountControl:1.2.840.113556.1.4.803:=2))

# Todos os usuários
(&(objectClass=user)(!(sAMAccountName=*$)))

# Usuários com email
(&(objectClass=user)(!(sAMAccountName=*$))(mail=*))
```

### Atributos Consultados
- `sAMAccountName` - Login do usuário
- `displayName` - Nome completo
- `title` - Cargo/função
- `mail` - E-mail corporativo
- `whenCreated` - Data de criação
- `userAccountControl` - Status da conta
- `lastLogon` - Último logon
- `lastLogonTimestamp` - Timestamp do logon
- `accountExpires` - Data de expiração

## 🛠️ Ferramentas Especializadas

### Busca_AD.py - Debug de Usuários
- Análise detalhada de dados crus do AD
- Conversão de datas e timestamps
- Validação de critérios de auditoria
- Análise de flags do userAccountControl

### Auditoria_2024.py - Relatório Completo
- Critérios específicos de auditoria
- Conversão robusta de datas
- Tratamento de casos especiais
- Relatório consolidado com estatísticas

### List_AD.py - Interface Principal
- Menu interativo amigável
- Múltiplas opções de relatório
- Geração automática de planilhas
- Navegação contínua entre relatórios

## 🐛 Resolução de Problemas

### Erro de Conexão
```
❌ ERRO: Não foi possível conectar ao Active Directory
```
**Soluções**:
- Verificar conectividade de rede
- Confirmar credenciais do usuário
- Verificar configurações de firewall
- Testar conectividade nas portas 389/636

### Erro de Permissões
```
❌ Erro ao buscar usuários: insuficientAccessRights
```
**Soluções**:
- Verificar permissões de leitura no AD
- Contatar administrador do domínio
- Executar com usuário com privilégios adequados

### Erro de Biblioteca
```
❌ Erro ao instalar openpyxl
```
**Soluções**:
- Instalar manualmente: `pip install openpyxl`
- Verificar conectividade com PyPI
- Usar ambiente virtual Python

### Erro de Conversão de Data
```
❌ Erro ao converter data de expiração
```
**Soluções**:
- Usar Busca_AD.py para debug do usuário específico
- Verificar formato dos timestamps no AD
- Confirmar configurações de timezone

## 📈 Logs e Debug

### Informações Exibidas
- Progresso da busca (páginas processadas)
- Estatísticas de usuários encontrados
- Critérios aplicados na filtragem
- Tempo de processamento
- Detalhes de conectividade
- Debug específico para casos especiais

### Exemplo de Log
```
🔍 Buscando usuários no Active Directory...
   Buscando página 1...
   ✓ Página 1: 1000 usuários | Total: 1000
   Buscando página 2...
   ✓ Página 2: 654 usuários | Total: 1654
   Busca concluída após 2 páginas
✅ Total de usuários encontrados: 1654

📊 ESTATÍSTICAS FINAIS:
   • Total processado: 1654 usuários
   • Incluídos na auditoria: 1420 usuários
   • Usuários ativos: 1200
   • Usuários inativos: 220
   • Excluídos: 234 usuários
```

## 🔍 Recursos Avançados

### Tratamento de Datas
- Conversão automática de timestamps do Windows
- Detecção de datas inválidas (1601-01-01)
- Tratamento de campos "nunca expira"
- Validação de períodos específicos

### Critérios de Auditoria Inteligentes
- Análise de contas com data de expiração
- Validação de último logon vs criação
- Tratamento de contas sem data válida
- Exclusão de contas antigas sem atividade

### Geração de Relatórios
- Nomes únicos com timestamp
- Abertura automática dos arquivos
- Formatação profissional
- Cabeçalhos informativos com critérios

## 🤝 Contribuições

Contribuições são bem-vindas! Por favor:

1. Faça um fork do projeto
2. Crie uma branch para sua feature (`git checkout -b feature/AmazingFeature`)
3. Commit suas mudanças (`git commit -m 'Add some AmazingFeature'`)
4. Push para a branch (`git push origin feature/AmazingFeature`)
5. Abra um Pull Request

## 📄 Licença

Este projeto é distribuído sob a licença MIT. Veja o arquivo `LICENSE` para mais detalhes.

## 📞 Suporte

Para suporte e questões técnicas:
- Abra uma issue no repositório
- Consulte a documentação do ldap3
- Verifique os logs de erro detalhados
- Use o Busca_AD.py para debug específico

## 📚 Referências

- [Documentação LDAP3](https://ldap3.readthedocs.io/)
- [OpenPyXL Documentation](https://openpyxl.readthedocs.io/)
- [Microsoft Active Directory LDAP](https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ldap/active-directory-ldap)

---

**Desenvolvido para Windows Server 2016+ com Active Directory**

**Versão Python recomendada: 3.8+**

**Última atualização: Julho 2025**