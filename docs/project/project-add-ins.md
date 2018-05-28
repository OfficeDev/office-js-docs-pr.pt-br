---
title: Suplementos do painel de tarefas para Project
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 07e64cca1d50f51e34f75f878044f2e02c9c4425
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="task-pane-add-ins-for-project"></a>Suplementos do painel de tarefas para Project

Tanto o Project Standard 2013 quanto o Project Professional 2013 incluem suporte para suplementos de painel de tarefas. Voc? pode executar suplementos de painel de tarefas comuns que foram desenvolvidos para o Word 2013 ou o Excel 2013. Voc? tamb?m pode desenvolver suplementos personalizados que manipulam eventos de sele??o no Project e integram tarefas, recursos, exibi??o e outros dados de n?vel de c?lula em um projeto com listas do SharePoint, Suplementos do SharePoint, Web Parts, servi?os Web e aplicativos corporativos.

> [!NOTE]
> O [download do SDK do Project 2013](https://www.microsoft.com/en-us/download/details.aspx?id=30435%20) inclui suplementos de exemplo que mostram como usar o modelo de objeto do suplemento no Project e como usar o servi?o OData para relatar os dados no Project Server 2013. Ao extrair e instalar o SDK, confira o subdiret?rio `\Samples\Apps\`.

Para ver uma introdu??o sobre os suplementos do Office, confira [Vis?o geral da plataforma de suplementos do Office](../overview/office-add-ins.md).

## <a name="add-in-scenarios-for-project"></a>Cen?rios de suplementos do Project

Os gerentes de projeto podem usar suplementos de painel de tarefas do Project para ajud?-los nas atividades de gerenciamento de projeto. Em vez de sair do Project e abrir outro aplicativo para procurar informa??es usadas com frequ?ncia, os gerentes de projeto podem acessar as informa??es diretamente no projeto. O conte?do de um suplemento de painel de tarefas pode ser contextual, baseado na tarefa selecionada, no recurso, no modo de exibi??o ou em outros dados em uma c?lula de um gr?fico de Gantt, no modo de exibi??o de uso da tarefa ou no modo de exibi??o de uso dos recursos.

> [!NOTE]
> Com o Project Professional 2013, ? poss?vel desenvolver suplementos de painel de tarefas que acessam instala??es locais do Project Server 2013, do Project Online e instala??es locais ou online do SharePoint 2013. O Project Standard 2013 n?o d? suporte ? integra??o direta com dados do Project Server ou listas de tarefas do SharePoint que s?o sincronizadas com o Project Server.

Cen?rios de suplementos do Project incluem o seguinte:

-  **Plano de projeto** Exibir dados de projetos relacionados que podem afetar o agendamento. Um suplemento de painel de tarefas pode integrar dados relevantes de outros projetos no Project Server 2013. Por exemplo, voc? pode exibir a cole??o de departamento de projetos e datas de marco ou exibir dados espec?ficos de outros projetos que s?o baseados em um campo personalizado selecionado.
    
-  **Gerenciamento de recursos** Exiba o pool de recursos completo no Project Server 2013 ou um subconjunto baseado em qualifica??es especificadas, incluindo a disponibilidade de dados de custo e recursos, para ajudar a selecionar recursos apropriados.
    
-  **Status e aprova??es** Use um aplicativo Web em um suplemento de painel de tarefas para atualizar ou exibir dados de um aplicativo de ERP (planejamento de recursos corporativos) externo, de um sistema de quadro de hor?rios ou de um aplicativo de contabilidade. Ou crie uma Web Part de aprova??o de status personalizada que pode ser usada no Project Web App e no Project Professional 2013.
    
-  **Comunica??o da equipe** Comunique-se com os membros da equipe e os recursos diretamente de um suplemento de painel de tarefas, dentro do contexto de um projeto. Ou mantenha um conjunto de anota??es contextuais para si mesmo facilmente enquanto trabalha em um projeto.
    
-  **Pacotes de trabalho** Pesquise tipos espec?ficos de modelos de projeto nas bibliotecas do SharePoint e em cole??es de modelos online. Por exemplo, encontre modelos para projetos de constru??o e adicione-os ? sua cole??o de modelos do Project.
    
-  **Itens relacionados** Exiba metadados, documentos e mensagens relacionadas a tarefas espec?ficas em um plano de projeto. Por exemplo, voc? pode usar o Project Professional 2013 para gerenciar um projeto que foi importado de uma lista de tarefas do SharePoint e ainda sincronizar a lista de tarefas com as altera??es no projeto. Um suplemento de painel de tarefas pode mostrar campos adicionais ou metadados que o Project n?o importou para tarefas na lista do SharePoint.
    
-  **Usar modelos de objeto do Project Server** Use o GUID de uma tarefa selecionada com m?todos na PSI (Project Server Interface) ou no CSOM (modelo de objeto do lado do cliente) do Project Server. Por exemplo, o aplicativo Web para um suplemento pode ler e atualizar os dados de status de uma tarefa e recurso selecionados ou integrar com um aplicativo de quadro de hor?rios externo.
    
-  **Obter dados de relat?rio** Use consultas LINQ, REST (Representational State Transfer) ou JavaScript para localizar informa??es relacionadas a uma tarefa ou recurso selecionado no servi?o OData para tabelas de relat?rio no Project Web App. Consultas que usam o servi?o OData podem ser feitas com instala??o online ou local do Project Server 2013.
    
    Por exemplo, confira [Criar um suplemento do Project que usa REST com um servi?o OData local do Project Server](../project/create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md).
    
## <a name="developing-project-add-ins"></a>Desenvolver suplementos do Project

A biblioteca JavaScript para suplementos do Project inclui extens?es do alias de namespace do **Office** que permitem que os desenvolvedores acessem propriedades de aplicativo do Project e tarefas, recursos e modos de exibi??o em um projeto. As extens?es de biblioteca JavaScript no arquivo Project-15.js s?o usadas em um suplemento do Project criado com o Visual Studio 2015. Office.js, Office.debug.js, Project-15.js, Project-15.debug.js e arquivos relacionados tamb?m s?o fornecidos no download do SDK do Project 2013.

Para criar um suplemento, voc? pode usar um editor de texto simples para criar uma p?gina da Web HTML e arquivos JavaScript relacionados, arquivos CSS e consultas REST. Al?m de uma p?gina HTML ou um aplicativo Web, um suplemento requer um arquivo de manifesto XML de configura??o. O Project pode usar um arquivo de manifesto que inclui um atributo **type** especificado como **TaskPaneExtension**. O arquivo de manifesto pode ser usado por v?rios aplicativos clientes do Office 2013, ou voc? pode criar um arquivo de manifesto que seja espec?fico para o Project 2013. Para saber mais, confira a se??o _No??es b?sicas sobre desenvolvimento_ em [Vis?o geral da plataforma de suplementos do Office](../overview/office-add-ins.md).

Para aplicativos personalizados complexos e depura??o mais f?cil, recomendamos que voc? use o Visual Studio 2015 no desenvolvimento de sites para suplementos. O Visual Studio 2015 inclui modelos para projetos de suplementos em que voc? pode escolher o tipo de suplemento (painel de tarefas, conte?do ou email) e o aplicativo host (Project, Word, Excel ou Outlook).  Para obter um exemplo que integra dados do Project Online, confira [Conectar um suplemento de painel de tarefas do Project ao PWA](http://blogs.msdn.com/b/project_programmability/archive/2012/11/02/connecting-a-project-task-pane-app-to-pwa.aspx) no blog Project Programmability do MSDN.

Quando voc? instala o download do SDK do Project 2013, o subdiret?rio `\Samples\Apps\` inclui os seguintes suplementos de exemplo:


-  **Pesquisa do Bing:** O arquivo de manifesto BingSearch.xml aponta para a p?gina de pesquisa do Bing para dispositivos m?veis. Como o aplicativo Web Bing j? existe na Internet, o suplemento Pesquisa do Bing n?o usa outros arquivos de c?digo-fonte ou o modelo de objeto de suplemento para o Project.
    
-  **Teste de modelo de objeto do Project:** O arquivo de manifesto JSOM_SimpleOMCalls.xml e o arquivo JSOM_Call.html s?o, juntos, um exemplo que testa o modelo de objeto e a funcionalidade do suplemento no Project 2013. O arquivo HTML faz refer?ncia ao arquivo JSOM_Sample.js, que tem fun??es JavaScript que usam o arquivo Office.js e o arquivo Project-15.js na funcionalidade principal. O download do SDK inclui todos os arquivos de c?digo-fonte necess?rios e o arquivo XML do manifesto para o suplemento Teste de modelo de objeto do Project. O desenvolvimento e a instala??o do exemplo Teste de modelo de objeto do Project est? descrito em [Criar seu primeiro suplemento de painel de tarefas para o Project 2013 usando um editor de texto](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).
    
-  **HelloProject_OData:** Essa ? uma solu??o do Visual Studio para o Project Professional 2013 que resume os dados do projeto ativo, como custo, trabalho e porcentagem conclu?da, e os compara com a m?dia de todos os projetos publicados na inst?ncia do Project Web App onde o projeto ativo est? armazenado. O desenvolvimento, a instala??o e o teste do exemplo, que usa o protocolo REST com o servi?o **ProjectData** no Project Web App, est?o descritos em [Criar um suplemento do Project que usa REST com um servi?o OData local do Project Server](../project/create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md).
    

### <a name="creating-an-add-in-manifest-file"></a>Criar um arquivo de manifesto do suplemento


O arquivo de manifesto especifica a URL do suplemento, a p?gina da Web ou aplicativo Web, o tipo de suplemento (painel de tarefas do Project), URLs opcionais de conte?do para outros idiomas e localidades, e outras propriedades.


### <a name="procedure-1-to-create-the-add-in-manifest-file-for-bing-search"></a>Procedimento 1. Para criar o arquivo de manifesto do suplemento para Pesquisa do Bing


- Crie um arquivo XML em um diret?rio local. O arquivo XML inclui o elemento **OfficeApp** e elementos filhos, que est?o descritos em [Manifesto XML dos suplementos do Office](../develop/add-in-manifests.md). Por exemplo, crie um arquivo denominado BingSearch.xml que cont?m o XML a seguir.
    
    ```XML
    <?xml version="1.0" encoding="utf-8"?>
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0" 
                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
              xsi:type="TaskPaneApp">
      <Id>1234-5678</Id>
      <Version>15.0</Version>
      <ProviderName>Microsoft</ProviderName>
      <DefaultLocale>en-us</DefaultLocale>
      <DisplayName DefaultValue="Bing Search">
      </DisplayName>
      <Description DefaultValue="Search selected data on Bing">
      </Description>
      <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg">
      </IconUrl>
      <Capabilities>
        <Capability Name="Project"/>
      </Capabilities>
      <DefaultSettings>
        <SourceLocation DefaultValue="http://m.bing.com">
        </SourceLocation>
      </DefaultSettings>
      <Permissions>ReadWriteDocument</Permissions>
    </OfficeApp>
    ```

- Abaixo est?o os elementos obrigat?rios no manifesto do suplemento:
  - No elemento **OfficeApp**, o atributo `xsi:type="TaskPaneApp"` especifica que o suplemento ? um tipo de painel de tarefas.
  - O elemento **Id** ? um UUID e precisa ser exclusivo.
  - O elemento **Version** ? a vers?o do suplemento. O elemento **ProviderName** ? o nome da empresa ou do desenvolvedor que fornece o suplemento. O elemento **DefaultLocale** especifica a localidade padr?o para as cadeias de caracteres no manifesto.
  - O elemento **DisplayName** ? o nome que mostra a lista suspensa **Suplemento do Painel de Tarefas** na guia **EXIBI??O** da faixa de op??es do Project 2013. O nome pode conter no m?ximo 32 caracteres.
  - O elemento **Description** cont?m a descri??o do suplemento para a localidade padr?o. O nome pode conter no m?ximo 2000 caracteres.
  - O elemento **Recursos** cont?m um ou mais elementos filhos **Capability** que especificam o aplicativo host.
  - O elemento **DefaultSettings** inclui o elemento **SourceLocation**, que especifica o caminho de um arquivo HTML em um compartilhamento de arquivo ou a URL de uma p?gina da Web que o suplemento usa. Um suplemento de painel de tarefas ignora os elementos **RequestedHeight** e **RequestedWidth**.
  - O elemento **IconUrl** ? opcional. Ele pode ser um ?cone em um compartilhamento de arquivo ou a URL de um ?cone em um aplicativo Web.
    
- (Opcional) Adicione elementos **Override** que t?m valores de outras localidades. Por exemplo, o manifesto a seguir fornece elementos **Override** para valores em franc?s de **DisplayName**, **Description**, **IconUrl** e **SourceLocation**.
    
    ```XML
    <?xml version="1.0" encoding="utf-8"?>
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0" 
                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
              xsi:type="TaskPaneApp">
      <Id>1234-5678</Id>
      <Version>15.0</Version>
      <ProviderName>Microsoft</ProviderName>
      <DefaultLocale>en-us</DefaultLocale>
      <DisplayName DefaultValue="Bing Search">
        <Override Locale="fr-fr" Value="Bing Search"/>
      </DisplayName>
      <Description DefaultValue="Search selected data on Bing">
        <Override Locale="fr-fr" Value="Search selected data on Bing"></Override>
      </Description>
      <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg">
        <Override Locale="fr-fr" Value="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg"/>
      </IconUrl>
      <Capabilities>
        <Capability Name="Project"/>
      </Capabilities>
      <DefaultSettings>
        <SourceLocation DefaultValue="http://m.bing.com">
          <Override Locale="fr-fr" Value="http://m.bing.com"/>
        </SourceLocation>
      </DefaultSettings>
      <Permissions>ReadWriteDocument</Permissions>
    </OfficeApp>
    ```


## <a name="installing-project-add-ins"></a>Instalar suplementos do Project


No Project 2013, ? poss?vel instalar suplementos como solu??es aut?nomas em um compartilhamento de arquivos ou em um cat?logo de suplementos particular. Tamb?m ? poss?vel avaliar e comprar suplementos no AppSource.

Pode haver v?rios arquivos XML do manifesto do suplemento e subdiret?rios em um compartilhamento de arquivos. Voc? pode adicionar ou remover locais e cat?logos do diret?rio do manifesto usando a guia **Cat?logos de Suplementos Confi?veis** na caixa de di?logo **Central de Confiabilidade** no Project 2013. Para mostrar um suplemento no Project, o elemento **SourceLocation** em um manifesto deve apontar para um site ou arquivo de origem HTML existente.


> [!NOTE]
> O Internet Explorer 9 ou posterior precisa estar instalado, mas n?o precisa ser o navegador padr?o. Os Suplementos do Office exigem componentes no Internet Explorer 9. O navegador padr?o pode ser o Internet Explorer 9, o Safari 5.0.6, o Firefox 5, o Chrome 13 ou uma vers?o mais recente de um desses navegadores.

No procedimento 2, o suplemento Pesquisa do Bing ? instalado no computador local onde o Project 2013 est? instalado. No entanto, como a infraestrutura do suplemento n?o usa caminhos de arquivo local diretamente, como `C:\Project\AppManifests`, voc? pode criar um compartilhamento de rede no computador local. Se preferir, voc? pode criar um compartilhamento de arquivos em um computador remoto.


### <a name="procedure-2-to-install-the-bing-search-add-in"></a>Procedimento 2. Para instalar o suplemento Pesquisa do Bing


1. Crie um diret?rio local para manifestos de suplemento. Por exemplo, crie o diret?rio `C:\Project\AppManifests`.
    
2. Compartilhe diret?rio `C:\Project\AppManifests` asAppManifests, para que o caminho de rede at? o compartilhamento de arquivo se torne `\\ServerName\AppManifests`.
    
3. Copie o arquivo de manifesto BingSearch.xml para o diret?rio `C:\Project\AppManifests`.
    
4. No Project 2013, abra caixa de di?logo **Op??es do Project**, escolha **Central de Confiabilidade** e escolha **Configura??es da Central de Confiabilidade**.
    
5. Na caixa de di?logo **Central de Confiabilidade**, no painel esquerdo, escolha **Cat?logos de Suplementos Confi?veis**.
    
6. No painel **Cat?logos de Suplementos Confi?veis** (confira a Figura 1), adicione o caminho `\\ServerName\AppManifests` ? caixa de texto **URL do Cat?logo**, escolha **Adicionar Cat?logo** e escolha **OK**.
    
    > [!NOTE]
    > A Figura 1 mostra dois compartilhamentos de arquivo e uma URL hipot?tica para um cat?logo particular na lista **Endere?os do Cat?logo Confi?vel**. Apenas um compartilhamento de arquivo pode ser o compartilhamento de arquivos padr?o, e apenas uma URL de cat?logo pode ser o cat?logo padr?o. Por exemplo, se voc? definir `\\Server2\AppManifests` como o padr?o, o Project limpar? a caixa de sele??o **Padr?o** para `\\ServerName\AppManifests`. Se voc? alterar a sele??o padr?o, escolha **Limpar** para remover suplementos instalados e reinicie o Project. Se voc? adicionar um suplemento ao compartilhamento de arquivo padr?o ou cat?logo do SharePoint enquanto o Project estiver aberto, reinicie o Project.

    *Figura 1. Usando a Central de Confiabilidade para adicionar cat?logos de manifestos de suplemento*

    ![Usar a Central de Confiabilidade para adicionar manifestos de aplicativo](../images/pj15-agave-overview-trust-centers.png)

7. Na faixa de op??es **Project**, escolha o menu suspenso **Suplementos do Office** e escolha **Ver Tudo**. Na caixa de di?logo **Inserir Suplemento**, escolha **PASTA COMPARTILHADA** (confira a Figura 2).
    
    *Figura 2. Iniciando um suplemento que est? em um compartilhamento de arquivos*

    ![Iniciar o aplicativo do Office que estiver em um compartilhamento de arquivos](../images/pj15-agave-overview-start-agave-apps.png)

8. Selecione o suplemento Pesquisa do Bing e escolha **Inserir**.
    
    O suplemento Pesquisa do Bing ? exibido em um painel de tarefas, como na Figura 3. Voc? pode redimensionar o painel de tarefas manualmente e usar o suplemento Pesquisa do Bing.

    *Figura 3. Usando o suplemento Pesquisa do Bing*

    ![Usar o aplicativo de Pesquisa do Bing](../images/pj15-agave-overview-bing-search.png)


## <a name="distributing-project-add-ins"></a>Distribuir suplementos do Project


? poss?vel distribuir suplementos usando um compartilhamento de arquivos, um cat?logo de suplementos em uma biblioteca do SharePoint ou o AppSource. Saiba mais em [Publicar seu suplemento do Office](../publish/publish.md).


## <a name="see-also"></a>Veja tamb?m

- [Vis?o geral da plataforma Suplementos do Office](../overview/office-add-ins.md)
- [Manifesto XML dos Suplementos do Office](../develop/add-in-manifests.md)
- [JavaScript API para Office](https://dev.office.com/reference/add-ins/javascript-api-for-office)
- [Criar seu primeiro suplemento de painel de tarefas para o Project 2013 usando um editor de texto](create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)
- [Criar um suplemento de Project que usa REST com um servi?o local do Project Server OData](create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md)
- [Conectar um suplemento de painel de tarefas do Project ao PWA](http://blogs.msdn.com/b/project_programmability/archive/2012/11/02/connecting-a-project-task-pane-app-to-pwa.aspx)
- [Download do SDK do Project 2013](https://www.microsoft.com/en-us/download/details.aspx?id=30435%20)
    
