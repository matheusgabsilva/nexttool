# NextTool

Ferramenta de TI da Next — baseada na ideia do [WinUtil](https://github.com/christitustech/winutil) de Chris Titus Tech, adaptada para o ambiente corporativo da Next.

## Execução via URL

```powershell
irm "https://raw.githubusercontent.com/matheusgabsilva/nexttool/main/nexttool.ps1" | iex
```

> Execute em um terminal PowerShell. A ferramenta solicita elevação automaticamente.

---

## Funcionalidades

### Instalações
| Software | Método |
|---|---|
| Adobe Acrobat Reader DC | winget |
| WinRAR | winget |
| AnyDesk | winget |
| TeamViewer | winget |
| Microsoft 365 | winget / ODT |
| Office 2021 Pro Plus | ODT (Office Deployment Tool) |
| Office 2016 Pro Plus | ODT (Office Deployment Tool) |

> Office 2021/2016 requer licença de volume (KMS/MAK) ou ativação manual.

### Tweaks
- Desativar Hibernação (`powercfg -h off`)
- Desativar Smart App Control (Windows 11)
- Atualizar drivers via Windows Update (PSWindowsUpdate)

### Manutenção
- Otimizar PC (temp, lixeira, limpeza de disco, flush DNS)
- Diagnóstico (processos, inicialização, Defender, Firewall, Event Viewer)
- SFC + DISM
- Flush DNS
- Acesso rápido à pasta de relatórios (`C:\Next-Relatorios`)

### Rede / Domínio
- Configurar DNS manualmente por adaptador
- Resetar DNS para DHCP
- Teste de conectividade (ping)
- Ver IP / ipconfig /all
- Ingresso em domínio AD com renomeação opcional

---

## Requisitos

- Windows 10 / 11
- PowerShell 5.1+
- Acesso à internet (para instalações e ODT)
- Winget instalado (App Installer — instalado automaticamente se ausente)

---

## Relatórios

Logs de execução são salvos em `C:\Next-Relatorios\`.

---

## Roadmap

- [ ] Integração com servidor de arquivos interno (`\\servidor\instaladores`)
- [ ] Perfil de instalação padrão por tipo de máquina
- [ ] Exportar relatório em PDF

---

*Desenvolvido pela equipe de TI da Next.*
