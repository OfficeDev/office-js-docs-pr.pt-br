
# <a name="requirements-for-running-office-add-ins"></a>Requisitos para a execução de Suplementos do Office


Este artigo descreve os requisitos de software e de dispositivo para execução de Suplementos do Office.

>
  **Observação:** Caso pretenda [publicar](../publish/publish.md) o suplemento na Office Store depois de criá-lo, verifique se você está em conformidade com as [Políticas de validação da Office Store](https://msdn.microsoft.com/en-us/library/jj220035.aspx). Por exemplo, para passar na validação, seu suplemento deve funcionar em todas as plataformas com suporte aos métodos que você definir (para mais informações, confira a [seção 4.12](https://msdn.microsoft.com/en-us/library/jj220035.aspx#Anchor_3) e a [Página de hospedagem e disponibilidade do suplemento do Office](https://dev.office.com/add-in-availability)).

Para uma visão avançada da compatibilidade atual dos suplementos do Office, confira a página [Disponibilidade de hosts e plataformas de suplementos do Office](http://dev.office.com/add-in-availability).

## <a name="server-requirements"></a>Requisitos de servidor

Para poder instalar e executar qualquer Suplemento do Office, primeiro você precisa implantar os arquivos de manifesto e de página da Web para a interface de usuário e o código de seu suplemento para os locais de servidor apropriados.

Para todos os tipos de suplementos (suplementos de conteúdo, do Outlook e de painel de tarefas, além dos comandos de suplemento), você precisa implantar seus arquivos de página da Web do suplemento em um servidor Web ou em um serviço de hospedagem da Web, como o [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md).


 >**Observação:**   ao desenvolver e depurar um suplemento no Visual Studio, este implanta e executa os arquivos de página da Web do suplemento localmente com o IIS Express, e não exige um servidor Web adicional. 

Para suplementos de conteúdo e de painel de tarefas, nos aplicativos host do Office compatíveis (aplicativos Web do Access, Word, Excel, PowerPoint ou Project) você também precisa de um [catálogo de suplementos](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) no SharePoint para carregar os arquivo de manifesto XML do suplemento.

Para testar e executar um suplemento do Outlook, a conta de email do Outlook do usuário deve residir no Exchange 2013 ou posterior, que está disponível pelo Office 365, Exchange Online ou por meio de uma instalação local. O usuário ou administrador instala os arquivos de manifesto para suplementos do Outlook nesse servidor.

 >**Observação:**   contas de email POP e IMAP no Outlook não são compatíveis com Suplementos do Office.




## <a name="client-requirements-windows-desktop-and-tablet"></a>Requisitos de cliente: Windows para área de trabalho e tablet

O software a seguir é necessário para o desenvolvimento de um Suplemento do Office para os clientes de área de trabalho ou da Web do Office compatíveis que são executados em dispositivos de área de trabalho, laptop ou tablet baseados em Windows:


- Para computadores de mesa com Windows x86 e x64, e tablets como o Surface Pro:
    - A versão de 32 ou de 64 bits do Office 2013 ou uma versão posterior, em execução no Windows 7 ou em uma versão posterior.
    - Excel 2013, Outlook 2013, PowerPoint 2013, Project Professional 2013, Project 2013 SP1, Word 2013 ou uma versão posterior do cliente do Office, se você estiver testando ou executando um Suplemento do Office especificamente para um desses clientes de área de trabalho do Office. É possível instalar clientes de área de trabalho do Office localmente ou por meio do recurso Clique para Executar no computador cliente.
    
        Se você tem uma assinatura válida do Office 365 e não tem acesso ao Office 2013, você pode baixá-lo por meio de um dos links CDN:
        
        - Office 2013 para empresas: [https://c2rsetup.officeapps.live.com/c2r/download.aspx?productReleaseID=O365BusinessRetail&platform=X86&language=en-us&version=O15GA&source=O15OLSO365](https://c2rsetup.officeapps.live.com/c2r/download.aspx?productReleaseID=O365BusinessRetail&platform=X86&language=en-us&version=O15GA&source=O15OLSO365) 
        - Office 2013 para uso doméstico: [https://c2rsetup.officeapps.live.com/c2r/download.aspx?productReleaseID=O365HomePremRetail&platform=X86&language=en-us&version=O15GA&source=O15OLSO365](https://c2rsetup.officeapps.live.com/c2r/download.aspx?productReleaseID=O365HomePremRetail&platform=X86&language=en-us&version=O15GA&source=O15OLSO365) 

- Internet Explorer 11 ou posterior, que deve estar instalado, mas não precisa ser o navegador padrão. Para oferecer suporte aos Suplementos do Office, o cliente do Office que atua como host usa os componentes do navegador que fazem parte do Internet Explorer 11 ou posterior.
- Um dos navegadores seguintes como o padrão: Internet Explorer 11 ou posterior, ou a versão mais recente do Microsoft Edge, Chrome, Firefox ou Safari (Mac OS).
- Um editor de HTML e JavaScript, como o Bloco de Notas, o [Visual Studio e Ferramentas de Desenvolvimento da Microsoft](https://www.visualstudio.com/features/office-tools-vs), ou uma ferramenta de desenvolvimento para Web de terceiros.

## <a name="client-requirements-os-x-desktop"></a>Requisitos de cliente: Computador com OS X

Outlook para Mac, que é distribuído como parte do Office 365 e oferece suporte a suplementos do Outlook. A execução de suplementos do Outlook no Outlook para Mac tem os mesmos requisitos que o próprio Outlook para Mac: o sistema operacional deve ser pelo menos o OS X v10.10 "Yosemite". Como o Outlook para Mac usa WebKit como um mecanismo de layout para processar as páginas do suplemento, não há qualquer dependência adicional de navegador.

Estas são as versões mínimas do cliente do Office para Mac que oferecem suporte a suplementos do Office:

- Word para Mac versão 15.18 (160109) 
- Excel para Mac versão 15.19 (160206) 
- PowerPoint para Mac versão 15.24 (160614)

## <a name="client-requirements-browser-support-for-office-online-web-clients-and-sharepoint"></a>Requisitos de cliente: Suporte do navegador para clientes da Web do Office Online e SharePoint

Qualquer navegador compatível com ECMAScript 5.1, HTML5 e CSS3, como o Internet Explorer 11 ou posterior, ou a versão mais recente do Microsoft Edge, do Chrome, do Firefox ou do Safari (Mac OS).


## <a name="client-requirements-non-windows-smartphone-and-tablet"></a>Requisitos do cliente: smartphones e tablets sem Windows

Especificamente para o OWA para Dispositivos e o Outlook Web App em execução em um navegador em smartphones e tablets sem Windows, os softwares a seguir são necessários para testar e executar suplementos do Outlook.


| Aplicativo host | Dispositivo | Sistema operacional | Conta do Exchange | Navegador móvel |
|:-----|:-----|:-----|:-----|:-----|
|OWA for Android|Smartphones Android. Tecnicamente, os dispositivos considerados "pequenos" ou "normais" pelo [SO Android](https://developer.android.com/guide/practices/screens_support.html).|Android 4.4 KitKat ou posterior|Atualização mais recente do Office 365 para empresas ou do Exchange Online|Suplemento nativo para Android, navegador não aplicável|
|OWA for iPad|iPad 2 ou posterior|iOS 6 ou posterior|Atualização mais recente do Office 365 para empresas ou do Exchange Online|Suplemento nativo para iOS, navegador não aplicável|
|OWA for iPhone|iPhone 4S ou posterior|iOS 6 ou posterior|Atualização mais recente do Office 365 para empresas ou do Exchange Online|Suplemento nativo para iOS, navegador não aplicável|
|Outlook Web App|iPhone 4 ou posterior, iPad 2 ou posterior, iPod Touch 4 ou posterior|iOS 5 ou posterior|Office 365, Exchange Online ou Exchange Server 2013 local ou posteriores|Safari|


## <a name="additional-resources"></a>Recursos adicionais

- [Visão geral da plataforma de suplementos do Office](../overview/office-add-ins.md)
- [Disponibilidade de host e plataforma para suplementos do Office](http://dev.office.com/add-in-availability)

