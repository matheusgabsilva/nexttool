# NextTool

Ferramenta de TI da Next — interface gráfica (WPF) para diagnóstico, manutenção e configuração de máquinas Windows no ambiente corporativo.

> **v4.1** — Reescrita com GUI WPF (dark theme), execução assíncrona e log em tempo real.

## Execução via URL

```powershell
irm nexttool.matheusgabsilva.digital | iex
```

> Execute em um terminal PowerShell como **Administrador**. A ferramenta solicita elevação automaticamente se necessário.

---

## Funcionalidades

### Diagnóstico
- Informações do sistema (CPU, RAM, disco, GPU, placa-mãe, BIOS)
- Status do Windows Defender, Firewall e TPM
- Últimos erros do Event Viewer (24h)
- Top 10 processos por uso de CPU/RAM
- Programas na inicialização

### Manutenção

#### 🧹 Limpeza
| Botão | Descrição |
|---|---|
| Otimizar PC | Limpa temporários, lixeira, `cleanmgr`, flush DNS |
| SFC + DISM | Verifica e repara arquivos do sistema |
| Verificar Disco (C:) | Agenda `chkdsk` na próxima inicialização |
| Limpar Cache WU | Para serviços, limpa `SoftwareDistribution\Download` e reinicia |
| Limpar Cache Miniaturas | Remove `thumbcache_*.db` do Explorer, mostra MB liberados |
| Limpar Credenciais | Remove todas as entradas salvas no Credential Manager |
| Limpar Cache Teams/Office | Limpa cache do Teams (novo e clássico), Office e Outlook |

#### 🌐 Rede
| Botão | Descrição |
|---|---|
| Flush DNS | `ipconfig /flushdns` |
| Reset Winsock/IP | Reseta pilha TCP/IP e Winsock |
| Renovar IP (DHCP) | `ipconfig /release` + `/renew`, exibe novo IP obtido |
| Resetar Proxy | Limpa proxy do WinInet (registro) e WinHTTP |
| Sincronizar Hora | `w32tm /resync /force`, exibe hora atualizada |

#### 🖨️ Impressão
| Botão | Descrição |
|---|---|
| Limpar Fila de Impressão | Para Spooler, limpa pasta `PRINTERS`, reinicia serviço |
| Reinstalar Impressoras | Remove impressoras e força redescoberta via PnP |

#### ⚙️ Sistema
| Botão | Descrição |
|---|---|
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

### Área de Trabalho
- Adicionar/remover ícones do sistema (Meu Computador, Rede, Lixeira, etc.)
- Gerenciar atalhos de aplicativos instalados na área de trabalho

### Rede
- Listar adaptadores de rede
- Configurar DNS manualmente por adaptador
- Resetar DNS para DHCP
- Teste de conectividade (ping 8.8.8.8 / 1.1.1.1 / google.com)
- `ipconfig /all`

### Domínio
- Ingressar máquina em domínio Active Directory
- Renomear o computador durante o ingresso

### Usuários
- Listar usuários locais
- Criar, remover e redefinir senha de usuários locais

---

## Requisitos

- Windows 10 / 11
- PowerShell 5.1+
- Acesso à internet (para Windows Update)

---

## Logs

Logs são exibidos em tempo real na interface e salvos em:

```
C:\Next-Relatorios\nexttool_<PC>_<yyyy-MM-dd_HH-mm-ss>.log
```

---

*Desenvolvido pela equipe de TI da Next.*
