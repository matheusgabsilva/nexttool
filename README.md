# NextTool

Ferramenta de TI da Next — baseada na ideia do [WinUtil](https://github.com/christitustech/winutil) de Chris Titus Tech, adaptada para o ambiente corporativo da Next.

> **v2.0** — Reescrita em CLI colorido (menu-driven), sem GUI. Leve, rápida, funciona em qualquer terminal PowerShell.

## Execução via URL

```powershell
irm "https://raw.githubusercontent.com/matheusgabsilva/nexttool/master/nexttool.ps1" | iex
```

> Execute em um terminal PowerShell. A ferramenta solicita elevação automaticamente.

---

## Funcionalidades

### [1] Instalações
| Software | Método |
|---|---|
| Adobe Acrobat Reader DC | winget |
| WinRAR | winget |
| AnyDesk | winget |
| TeamViewer | winget |
| Microsoft 365 | ODT |
| Office 2021 Pro Plus VL | ODT |
| Office 2016 Pro Plus VL | ODT |

Opção **[8]** permite instalar múltiplos de uma vez (ex: `1,2,3`).

> Office 2021/2016 requer licença de volume (KMS/MAK) ou ativação manual.

### [2] Tweaks
- Desativar Hibernação (`powercfg -h off`)
- Desativar Smart App Control (Windows 11)
- Atualizar drivers via Windows Update (PSWindowsUpdate)
- Aplicar todos em sequência

### [3] Manutenção
- Otimizar PC (temp, lixeira, `cleanmgr /sagerun:64`, flush DNS, relatório RAM/disco)
- Diagnóstico (top 10 processos, inicialização, Defender, Firewall, Event Viewer 24h)
- SFC + DISM RestoreHealth
- Flush DNS
- Abrir pasta de relatórios

### [4] Rede / Domínio
- Listar adaptadores
- Configurar DNS manualmente por adaptador
- Resetar DNS para DHCP
- Teste de conectividade (ping em 8.8.8.8 / 1.1.1.1 / google.com)
- `ipconfig /all`
- Ingresso em domínio AD com renomeação opcional

---

## Requisitos

- Windows 10 / 11
- PowerShell 5.1+
- Acesso à internet (para instalações e ODT)
- Winget instalado (App Installer — instalado automaticamente se ausente)

---

## Relatórios

Logs coloridos são exibidos em tela e persistidos em:

```
C:\Next-Relatorios\nexttool_<PC>_<yyyy-MM-dd_HH-mm-ss>.log
```

---

## Roadmap

- [ ] Integração com servidor de arquivos interno (`\\servidor\instaladores`)
- [ ] Perfil de instalação padrão por tipo de máquina
- [ ] Exportar relatório em PDF

---

*Desenvolvido pela equipe de TI da Next.*
