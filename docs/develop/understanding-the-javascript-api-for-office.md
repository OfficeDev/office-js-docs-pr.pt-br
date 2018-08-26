---
title: Noções básicas da API JavaScript para Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: a9e1e26d4ba94a933ecb98250c19afee90750f5d
ms.sourcegitcommit: 28fc652bded31205e393df9dec3a9dedb4169d78
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/23/2018
ms.locfileid: "22928032"
---
# <a name="understanding-the-javascript-api-for-office"></a>Noções básicas da API JavaScript para Office

Este artigo fornece informações sobre a API JavaScript para Office e como usá-la. Para referenciar as informações, consulte [API JavaScript para Office](https://dev.office.com/reference/add-ins/javascript-api-for-office). Para obter informações sobre como atualizar os arquivos de projeto do Visual Studio para a versão mais recente da API JavaScript para Office, consulte [Atualizar a versão da API JavaScript para Office e arquivos de esquema do manifesto](update-your-javascript-api-for-office-and-manifest-schema-version.md).

> [!NOTE]
> Caso pretenda [publicar](../publish/publish.md) o suplemento na experiência do Office depois de criá-lo, verifique se você está em conformidade com as [Políticas de validação do AppSource](https://docs.microsoft.com/office/dev/store/validation-policies). Por exemplo, para passar na validação, seu suplemento deve funcionar em todas as plataformas com suporte aos métodos que você definir (para mais informações, confira a [seção 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) e a [Página de hospedagem e disponibilidade de suplementos do Office](../overview/office-add-in-availability.md)). 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a>Fazer referência à biblioteca da API JavaScript para Office no suplemento

A biblioteca da [API JavaScript para Office](https://dev.office.com/reference/add-ins/javascript-api-for-office) consiste no arquivo Office.js e nos arquivos .js específicos do aplicativo de host associado, como Excel-15.js e Outlook-15.js. O método mais simples de fazer referência à API é usando nossa CDN e adicionando o seguinte `<script>` à marca `<head>` da sua página:  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

Isso baixará e colocará os arquivos da API JavaScript para Office em cache quando o suplemento for carregado pela primeira vez a fim de garantir que o suplemento esteja usando a implementação mais recente do Office.js e de seus arquivos associados na versão especificada.

Para saber mais sobre a CDN do Office.js, incluindo como é feito o controle de versão e como lidar com a compatibilidade com versões anteriores, veja [Fazer referência à biblioteca da API JavaScript para Office a partir da sua rede de distribuição de conteúdo (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).

## <a name="initializing-your-add-in"></a>Iniciar o suplemento

**Aplica-se a:** todos os tipos de suplementos

O Office.js fornece um evento de inicialização que é acionado quando a API está totalmente carregada e pronta para começar a interação com o usuário. Você pode usar o manipulador de eventos **initialize** para implementar cenários comuns de inicialização de suplementos, como solicitar que o usuário selecione algumas células no Excel e, em seguida, insira um gráfico gerado a partir desses valores selecionados. Você também pode usar o manipulador de eventos de inicialização para inicializar outras lógicas personalizadas do suplemento, como estabelecer associações, solicitar valores padrão de configuração do suplemento e assim por diante.

No mínimo, o evento de inicialização se pareceria com o exemplo a seguir:     

```js
Office.initialize = function () { };
```
Se você estiver usando estruturas JavaScript adicionais que incluem seus próprios manipuladores de inicialização ou testes, esses devem ser colocados dentro do evento Office.initialize. Por exemplo, a função [JQuery](https://jquery.com) `$(document).ready()` seria referenciada da seguinte maneira:

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {        
        // The document is ready
    });
  };
```

Todas as páginas dentro de Suplementos do Office são necessárias para atribuir um manipulador de eventos ao evento de inicialização, **Office.initialize**. Se você não incluir um manipulador de eventos, o suplemento poderá gerar um erro ao iniciar. Além disso, se um usuário tentar usar o suplemento com um cliente Web do Office Online, como o Excel Online, o PowerPoint Online ou o Outlook Web App, ele não funcionará. Se você não precisar de nenhum código de inicialização, então, o corpo da função atribuída a **Office.initialize** poderá ficar vazio, como no primeiro exemplo acima.

Para obter mais detalhes sobre a sequência de eventos na inicialização do suplemento, consulte [Carregar o DOM e o ambiente de execução](loading-the-dom-and-runtime-environment.md).

#### <a name="initialization-reason"></a>Motivo da inicialização
Para suplementos de conteúdo e de painel de tarefas, o Office.initialize fornece um parâmetro _reason_ adicional. Esse parâmetro pode ser usado para determinar como um suplemento foi adicionado ao documento atual. Você pode usar isso para fornecer lógica diferente para quando um suplemento pela primeira vez em comparação com quando já existia dentro do documento. 

```js
Office.initialize = function (reason) {
    $(document).ready(function () {
        switch (reason) {
            case 'inserted': console.log('The add-in was just inserted.');
            case 'documentOpened': console.log('The add-in is already part of the document.');
        }
    });
 };
```
Para obter mais informações, confira [Evento Office.initialize](https://dev.office.com/reference/add-ins/shared/office.initialize) e [Enumeração InitializationReason](https://dev.office.com/reference/add-ins/shared/initializationreason-enumeration). 

## <a name="office-javascript-api-object-model"></a>Modelo de objeto da API JavaScript para Office

Uma vez inicializado, o suplemento pode interagir com o host (ex. Excel, Outlook). A página do [Modelo de Objeto da API Javascript do Office](office-javascript-api-object-model.md) tem mais detalhes sobre modelos de uso específicos. Também há documentação de referência detalhada para [APIs compartilhadas](https://dev.office.com/reference/add-ins/javascript-api-for-office) e hosts específicos.

## <a name="api-support-matrix"></a>Matriz de suporte da API


Esta tabela resume a API e os recursos compatíveis com os tipos de suplemento (conteúdo, painel de tarefas e Outlook) e os aplicativos do Office que podem hospedá-los quando o usuário especifica os aplicativos hospedados pelo Office compatíveis com o suplemento usando o [esquema 1.1 do manifesto de suplementos e recursos compatíveis com a v1.1 da API JavaScript para Office](update-your-javascript-api-for-office-and-manifest-schema-version.md).


|||||||||
|:-----|:-----|:-----:|:-----:|:-----:|:-----:|:-----:|:-----:|
||**Nome do host**|Banco de dados|Pasta de trabalho|Caixa de correio|Apresentação|Documento|Projeto|
||**Aplicativos host** **compatíveis**|Aplicativos Web do Access|Excel,<br/>Excel Online|Outlook,<br/>Outlook Web App,<br/>OWA para dispositivos|PowerPoint,<br/>PowerPoint Online|Word|Projeto|
|**Tipos de suplemento com suporte**|Conteúdo|S|S||S|||
||Painel de tarefas||S||S|S|S|
||Outlook|||S||||
|**Recursos da API compatíveis**|Ler/gravar texto||S||S|S|S<br/>(Somente leitura)|
||Ler/gravar matriz||S|||S||
||Ler/gravar tabela||S|||S||
||Ler/gravar HTML|||||S||
||Leitura/gravação<br/>Office Open XML|||||S||
||Ler propriedades de tarefa, recurso, modo de exibição e campo||||||S|
||Eventos alterados pela seleção||S|||S||
||Obter documento inteiro||||S|S||
||Associações e eventos de associação|S<br/>(Somente vinculações de tabela totais e parciais)|S|||S||
||Ler/gravar partes XML personalizadas|||||S||
||Persistir dados de estado de suplemento (configurações)|S<br/>(Por suplemento do host)|S<br/>(Por documento)|S<br/>(Por caixa de correio)|S<br/>(Por documento)|S<br/>(Por documento)||
||Eventos alterados pelas configurações|S|S||S|S||
||Obter o modo de exibição ativo<br/>e visualizar eventos alterados||||S|||
||Navegar para locais<br/>no documento||S||S|S||
||Ativar contextualmente<br/>usando regras e RegEx|||S||||
||Ler propriedades do item|||S||||
||Ler perfil de usuário|||S||||
||Obter anexos|||S||||
||Obter o token de identidade do usuário|||S||||
||Chamar os serviços Web do Exchange|||S||||
