# NextTool

Ferramenta de TI da Next — interface gráfica (WPF) para diagnóstico, manutenção e configuração de máquinas Windows no ambiente corporativo.

> **v4.0** — Reescrita com GUI WPF (dark theme), execução assíncrona e log em tempo real.

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
- Limpeza de arquivos temporários e lixeira
- SFC + DISM RestoreHealth
- Flush DNS
- Windows Update (via PSWindowsUpdate)
- Limpeza de disco (`cleanmgr`)

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
| Desativar Suspender (Sleep) | Impede que o PC entre em suspensão |
| Desativar Desligamento de Tela | Mantém a tela sempre ligada |
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
