---
title: Implante e instale suplementos do Outlook para teste
description: Crie um arquivo de manifesto, implante o arquivo de interface do usuário suplemento em um servidor web, instale o suplemento na caixa de correio e teste o suplemento.
ms.date: 07/08/2021
localization_priority: Priority
ms.openlocfilehash: 7a8c6a22d053f58b38e6dbb432c2ad151c0fee94f3a512eea0611e28ab3595f6
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57086238"
---
# <a name="deploy-and-install-outlook-add-ins-for-testing"></a>Implante e instale suplementos do Outlook para teste

Como parte do processo de desenvolvimento de um suplemento do Outlook, você provavelmente se verá implantando e instalando iterativamente o suplemento para teste, o que envolve as etapas a seguir.

1. Criação de um arquivo de manifesto que descreve o suplemento.
1. Implantação dos arquivos da interface do usuário em um servidor Web.
1. Instalação do suplemento em sua caixa de correio.
1. Teste do suplemento, fazendo as alterações apropriadas na interface de usuário ou nos arquivos de manifesto e repetindo as etapas 2 e 3 para testar as alterações.

> [!NOTE]
> [Os painéis personalizados foram preteridos](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/); portanto, certifique-se de estar usando um [ponto de extensão de suplemento com suporte](outlook-add-ins-overview.md#extension-points).

## <a name="create-a-manifest-file-for-the-add-in"></a>Criar um arquivo de manifesto para o suplemento

Cada suplemento é descrito por um manifesto XML, um documento que fornece as informações do servidor sobre o suplemento, fornece informações sobre o suplemento para o usuário e identifica o local da interface do arquivo HTML de interface do usuário do suplemento. Você pode armazenar o manifesto em uma pasta ou servidor local, desde que o local possa ser acessado pelo servidor Exchange da caixa de correio que você está testando. Vamos pressupor que você armazena seu manifesto em uma pasta local. Para obter informações sobre como criar um arquivo de manifesto, confira [Manifestos de suplementos do Outlook](manifests.md).

## <a name="deploy-an-add-in-to-a-web-server"></a>Implantar um suplemento em um servidor Web

Você pode usar HTML e JavaScript para criar o suplemento. Os arquivos de origem resultantes são armazenados em um servidor Web que pode ser acessado pelo servidor Exchange que hospeda o suplemento. Depois de implantar inicialmente os arquivos de origem para o suplemento, você pode atualizar a interface do usuário e o comportamento dele substituindo os arquivos HTML ou JavaScript armazenados no servidor Web por uma nova versão do arquivo HTML.

## <a name="install-the-add-in"></a>Instalar o suplemento

Depois de preparar o arquivo de manifesto do suplemento e implantar a interface de usuário do suplemento em um servidor Web que possa ser acessado, é possível realizar o sideload do suplemento para uma caixa de correio em um servidor Exchange usando um cliente do Outlook ou instalar o suplemento executando cmdlets remotos do Windows PowerShell.

### <a name="sideload-the-add-in"></a>Realizar o sideload do suplemento

Você pode instalar um suplemento se sua caixa de correio está no Exchange Online, no Exchange 2013 ou em uma versão posterior. Os suplementos de sideload exigem no mínimo a função **Meus Aplicativos Personalizados** do seu Exchange Server. Para testar seu suplemento ou instalar suplementos em geral especificando uma URL ou um nome de arquivo de manifesto do suplemento, é preciso solicitar que o administrador do Exchange forneça as permissões necessárias.

O administrador do Exchange pode executar o cmdlet do PowerShell a seguir para atribuir as permissões necessárias a um único usuário. Neste exemplo, `wendyri` é o alias de email do usuário.

```powershell
New-ManagementRoleAssignment -Role "My Custom Apps" -User "wendyri"
```

Se necessário, o administrador poderá executar o seguinte cmdlet para atribuir a vários usuários as permissões semelhantes necessárias.

```powershell
$users = Get-Mailbox *$users | ForEach-Object { New-ManagementRoleAssignment -Role "My Custom Apps" -User $_.Alias}
```

Para saber mais sobre a função Meus Suplementos Personalizados, confira [Função Meus Suplementos Personalizados](/exchange/my-custom-apps-role-exchange-2013-help).

O uso do Microsoft 365 ou do Visual Studio para desenvolver suplementos atribui a você a função de administrador da organização, o que permite que você instale suplementos por arquivo ou URL no EAC, ou por cmdlets do Powershell.

### <a name="install-an-add-in-by-using-remote-powershell"></a>Instalar um suplemento usando o PowerShell remoto

Depois de criar uma sessão remota do Windows PowerShell em seu servidor Exchange, você pode instalar um suplemento do Outlook usando o cmdlet `New-App` com o comando do PowerShell a seguir.

```powershell
New-App -URL:"http://<fully-qualified URL">
```

A URL totalmente qualificada é o local do arquivo de manifesto do suplemento que você preparou para seu suplemento.

Você pode usar os seguintes cmdlets adicionais do Windows PowerShell para gerenciar os suplementos de uma caixa de correio.

- `Get-App` – Lista os suplementos que estão habilitados para uma caixa de correio.
- `Set-App` – Habilita ou desabilita um suplemento em uma caixa de correio.
- `Remove-App` – Remove um suplemento instalado anteriormente de um servidor Exchange.

## <a name="client-versions"></a>Versões de cliente

A decisão de quais versões de cliente do Outlook testar depende dos seus requisitos de desenvolvimento.

- Se você estiver desenvolvendo um suplemento para uso privado ou apenas para membros da sua organização, é importante testar as versões do Outlook usadas pela sua empresa. Lembre-se que alguns usuários podem usar o Outlook na Web, portanto, testar as versões para o navegador padrão da sua empresa também é importante.

- Se você estiver desenvolvendo um suplemento para listar no [AppSource](https://appsource.microsoft.com), deverá testar as versões necessárias conforme especificado nas [Políticas de certificação do Commercial Marketplace 1120.3](/legal/marketplace/certification-policies#11203-functionality). Isso inclui:
  - A versão mais recente do Outlook no Windows e a versão anterior à mais recente.
  - A versão mais recente do Outlook no Mac.
  - A versão mais recente do Outlook no iOS e Android (se o suplemento for [compatível com mobilidade](add-mobile-support.md)).
  - As versões do navegador especificadas na política de validação do mercado comercial 1120.3

> [!NOTE]
> Se seu suplemento não for compatível com um dos clientes acima devido a uma [solicitação de um conjunto de exigências da API](apis.md) que o cliente não dá suporte, esse cliente será removido da lista de clientes exigidos.

## <a name="outlook-on-the-web-and-exchange-server-versions"></a>Versões do Outlook na Web e do Exchange Server

Os usuários de contas do Microsoft 365 e de consumidor veem a versão moderna da interface do usuário ao acessar o Outlook na Web e não veem mais a versão clássica que foi substituída. No entanto, os servidores locais do Exchange continuam oferecendo suporte ao Outlook na Web clássico. Portanto, durante o processo de validação, seu envio poderá receber um aviso de que o suplemento não é compatível com o Outlook na Web clássico. Nesse caso, considere testar o suplemento em um ambiente do Exchange local. Esse aviso não bloqueará seu envio ao AppSource, mas seus clientes poderão ter uma experiência abaixo do ideal, caso usem o Outlook na Web em um ambiente do Exchange local.

Para atenuar isso, é recomendável que se faça o teste do suplemento no Outlook na Web conectado ao seu próprio ambiente Exchange local. Para saber mais, confira as orientações sobre como [Estabelecer um ambiente de teste do Exchange 2016 ou do Exchange 2019](/Exchange/plan-and-deploy/plan-and-deploy?view=exchserver-2019&preserve-view=true#establish-an-exchange-2016-or-exchange-2019-test-environment) e como gerenciar o [Outlook na Web no Exchange Server](/exchange/clients/outlook-on-the-web/outlook-on-the-web?view=exchserver-2019&preserve-view=true).

Como alternativa, você pode optar por pagar e usar um serviço que hospeda e gerencia servidores Exchange locais. Algumas opções são:

- [Rackspace](https://www.rackspace.com/email-hosting/exchange-server)
- [Hostway](https://hostway.com/microsoft-exchange/)

Além disso, se você não deseja que seus suplementos estejam disponíveis para usuários conectados ao Exchange local, é possível definir o [conjunto de requisitos](../reference/requirement-sets/outlook-api-requirement-sets.md#exchange-server-support) no manifesto de suplemento como 1.6 ou superior. Esses suplementos não serão testados nem validados na interface do usuário do Outlook na Web clássico.

## <a name="see-also"></a>Confira também

- [Solucionar erros de usuários com Suplementos do Office](../testing/testing-and-troubleshooting.md)
