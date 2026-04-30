# NextTool

Ferramenta de TI da Next — interface gráfica (WPF) para diagnóstico, manutenção e configuração de máquinas Windows no ambiente corporativo.

> **v4.2** — Modo Usuário / ADM com senha, painel simplificado para usuários finais, guias de uso integrados.

## Execução via URL

```powershell
irm nexttool.matheusgabsilva.digital | iex
```

> Execute em um terminal PowerShell como **Administrador**. A ferramenta solicita elevação automaticamente se necessário.

---

## Modos de acesso

A ferramenta inicia sempre no **Modo Usuário** — sem senha, com funções seguras para o dia a dia. Para acessar as funcionalidades avançadas, clique em **Entrar ADM** e informe a senha.

| | Modo Usuário | Modo ADM |
|---|---|---|
| Acesso | Direto, sem senha | Senha protegida |
| Sistema | Info do PC, IP, RAM, domínio | Diagnóstico completo + relatório HTML |
| Limpeza | Temporários, Teams, impressão, DNS, IP, hora | Todas as funções + SFC, DISM, WU |
| Rede | Ping, Tracert, Testar Conexão | Configurar DNS, ingressar em domínio |
| Armazenamento | Discos + analisar pastas | Completo |
| Tweaks / Usuários | — | Disponível |

**Senha padrão:** configurada na variável `$script:SENHA_PADRAO` no topo do script.  
A senha é armazenada como hash SHA256 em `HKLM:\SOFTWARE\NextTool\AdminHash`.

---

## Funcionalidades — Modo ADM

### Sistema
- Informações completas de hardware (CPU, RAM, disco, GPU, placa-mãe, BIOS)
- IP, Gateway, DNS, adaptador de rede, domínio e licença Windows visíveis na tela inicial
- Status do Defender, Firewall, UAC, TPM e Secure Boot
- Últimos erros críticos do Event Viewer (24h)
- Top 10 processos por uso de RAM
- Programas na inicialização
- Exportar relatório HTML completo (abre no navegador)
- Histórico de sessões anteriores

### Manutenção

#### 🧹 Limpeza
| Botão | Descrição |
|---|---|
| Otimizar PC | Limpa temporários, lixeira, `cleanmgr`, flush DNS |
| SFC + DISM | Verifica e repara arquivos do sistema |
| Verificar Disco (C:) | Agenda `chkdsk` na próxima inicialização |
| Limpar Cache WU | Para serviços, limpa `SoftwareDistribution\Download` e reinicia |
| Limpar Cache Miniaturas | Remove `thumbcache_*.db` do Explorer |
| Limpar Credenciais | Remove entradas salvas no Credential Manager |
| Limpar Cache Teams/Office | Limpa cache do Teams (novo e clássico), Office e Outlook |

#### 🌐 Rede
| Botão | Descrição |
|---|---|
| Flush DNS | `ipconfig /flushdns` |
| Reset Winsock/IP | Reseta pilha TCP/IP e Winsock |
| Renovar IP (DHCP) | `ipconfig /release` + `/renew`, exibe novo IP |
| Resetar Proxy | Limpa proxy do WinInet e WinHTTP |
| Sincronizar Hora | `w32tm /resync /force` |

#### 🖨️ Impressão
| Botão | Descrição |
|---|---|
| Limpar Fila de Impressão | Para Spooler, limpa pasta `PRINTERS`, reinicia |
| Reinstalar Impressoras | Remove e força redescoberta via PnP |

#### ⚙️ Sistema
| Botão | Descrição |
|---|---|
| Verificar Updates | Lista atualizações pendentes sem instalar |
| Status de Serviços | Exibe status de 10 serviços críticos |
| Reiniciar Serviço | Reinicia qualquer serviço pelo nome |
| Limpar Perfil Corrompido | Remove perfis `.bak` e `TEMP` do registro |
| Desinstalar App (winget) | Desinstala aplicativo pelo nome via winget |
| gpupdate /force | Força atualização de políticas de grupo |
| Reiniciar Explorer | Encerra e reinicia o Windows Explorer |

### Tweaks
| Tweak | Descrição |
|---|---|
| Desativar Telemetria | Desliga coleta de dados da Microsoft |
| Histórico de Atividades | Desativa rastreamento de atividades |
| Rastreamento de Localização | Desliga GPS/localização do sistema |
| Extensões de Arquivo | Exibe extensões no Explorador |
| Arquivos Ocultos | Exibe arquivos e pastas ocultos |
| NumLock na inicialização | Ativa NumLock automaticamente |
| Encerrar Tarefa (barra) | Habilita "Encerrar Tarefa" no menu da barra |
| Serviços desnecessários | Desativa SysMain, DiagTrack, WSearch |
| Desativar Hibernação | `powercfg -h off` |
| Smart App Control | Desativa (Windows 11) |
| Desativar Suspender (Sleep) | Impede que o PC entre em suspensão (AC e DC) |
| Desativar Desligamento de Tela | Mantém a tela sempre ligada (AC e DC) |
| Performance Máxima | Ativa plano de energia Ultimate Performance |
| Tema Escuro | Aplica dark mode no sistema e apps |
| Desativar Widgets | Remove widgets da barra de tarefas |
| Logon Detalhado | Exibe mensagens detalhadas na inicialização |

> Botão **Padrão Next** marca automaticamente o conjunto de tweaks recomendado pela equipe.

### Rede / Domínio
- Ping, Tracert e Testar Conectividade com diagnóstico visual
- Listar adaptadores de rede
- Configurar DNS manualmente por adaptador
- Resetar DNS para DHCP / Renovar DHCP
- `ipconfig /all`
- Verificar status de domínio AD
- Ingressar máquina em domínio Active Directory
- Renomear o computador durante o ingresso

### Usuários
- Listar usuários locais com status, admin e última senha
- Criar, remover, ativar/desativar e redefinir senha
- Desbloquear conta bloqueada por tentativas erradas
- Adicionar usuário ao grupo Administradores

### Área de Trabalho
- Adicionar/remover ícones do sistema (Meu Computador, Rede, Lixeira, etc.)
- Gerenciar atalhos de aplicativos instalados (com busca e seleção múltipla)

### Armazenamento
- Barras de uso por disco com indicador de espaço
- Analisar pasta: lista subpastas e maiores arquivos por tamanho
- Duplo clique abre a pasta/arquivo diretamente no Explorer

---

## Como enviar logs ao suporte

1. Execute o **Diagnóstico Completo** na aba Sistema
2. Clique em **Exportar** no rodapé da tela
3. O arquivo será salvo em `C:\Next-Relatorios\` com o nome:

```
Next_Suporte_<NOMEDOPC>_<yyyy-MM-dd_HH-mm-ss>.txt
```

4. Encaminhe o arquivo ao suporte por e-mail ou WhatsApp

---

## Requisitos

- Windows 10 / 11
- PowerShell 5.1+
- Acesso à internet (para Windows Update e winget)

---

## Logs de sessão

Logs são salvos automaticamente em:

```
C:\Next-Relatorios\nexttool_<PC>_<yyyy-MM-dd_HH-mm-ss>.log
```

---

*Desenvolvido pela equipe de TI da Next.*
