
# <a name="understanding-the-javascript-api-for-office"></a>Noções básicas da API JavaScript para Office



Este artigo fornece informações sobre a API JavaScript para Office e como usá-la. Para referenciar as informações, consulte [API JavaScript para Office](http://dev.office.com/reference/add-ins/javascript-api-for-office). Para obter informações sobre como atualizar os arquivos de projeto do Visual Studio para a versão mais recente da API JavaScript para Office, consulte [Atualizar a versão da API JavaScript para Office e arquivos de esquema do manifesto](../../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md).

>
  **Observação:** Caso pretenda [publicar](../publish/publish.md) o suplemento na Office Store depois de criá-lo, verifique se você está em conformidade com as [Políticas de validação da Office Store](https://msdn.microsoft.com/en-us/library/jj220035.aspx). Por exemplo, para passar na validação, seu suplemento deve funcionar em todas as plataformas com suporte aos métodos que você definir (para mais informações, confira a [seção 4.12](https://msdn.microsoft.com/en-us/library/jj220035.aspx#Anchor_3) e a [Página de hospedagem e disponibilidade do suplemento do Office](https://dev.office.com/add-in-availability)).

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a>Fazer referência à biblioteca da API JavaScript para Office no suplemento

A biblioteca da [API JavaScript para Office](http://dev.office.com/reference/add-ins/javascript-api-for-office) consiste no arquivo Office.js e nos arquivos .js específicos do aplicativo de host associado, como Excel-15.js e Outlook-15.js. O método mais simples de fazer referência à API é usando nossa CDN e adicionando o seguinte `<script>` à marca `<head>` da sua página:  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

Isso baixará e colocará os arquivos da API JavaScript para Office em cache quando o suplemento for carregado pela primeira vez a fim de garantir que o suplemento esteja usando a implementação mais recente do Office.js e de seus arquivos associados na versão especificada.

Para saber mais sobre a CDN do Office.js, incluindo como é feito o controle de versão e como lidar com a compatibilidade com versões anteriores, consulte [Fazer referência à biblioteca da API JavaScript para Office de sua CDN (rede de distribuição de conteúdo)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).

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

Para obter mais detalhes sobre a sequência de eventos na inicialização do suplemento, consulte [Carregar o DOM e o ambiente de execução](../../docs/develop/loading-the-dom-and-runtime-environment.md).

#### <a name="initialization-reason"></a>Motivo da inicialização
Para suplementos de conteúdo e de painel de tarefas, o Office.initialize fornece um parâmetro _reason_ adicional. Esse parâmetro pode ser usado para determinar como um suplemento foi adicionado ao documento atual. Você pode usar isso para fornecer lógica diferente para quando um suplemento pela primeira vez em comparação com quando já existia dentro do documento. 

```js
Office.initialize = function (reason) {
    $(document).ready(function () {
      switch (reason) {
        case 'inserted': console.log('The add-in was just inserted.');
        case 'documentOpened': console.log('The add-in is already part of the document.');
    }
}
```
Para obter mais informações, confira [Evento Office.initialize](http://dev.office.com/reference/add-ins/shared/office.initialize) e [Enumeração InitializationReason](http://dev.office.com/reference/add-ins/shared/initializationreason-enumeration) 

## <a name="context-object"></a>Objeto de contexto

 **Aplica-se a:** todos os tipos de suplementos

Quando um suplemento é iniciado, ele possui diversos objetos diferentes com os quais pode interagir no ambiente de tempo de execução. O contexto do tempo de execução do suplemento é refletido na API por meio do objeto [Context](http://dev.office.com/reference/add-ins/shared/office.context). **Context** é o principal objeto e fornece acesso aos objetos mais importantes da API, como [Document](http://dev.office.com/reference/add-ins/shared/document) e [Mailbox](http://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) que, por sua vez, fornecem acesso ao conteúdo do documento e da caixa de correio.

Por exemplo, nos suplementos do painel de tarefas e de conteúdo, é possível usar a propriedade [documento](http://dev.office.com/reference/add-ins/shared/office.context.document) do objeto **Context** para acessar as propriedades e os métodos do objeto **Document**. Isso permite interagir com o conteúdo de documentos do Word, planilhas do Excel ou tarefas do Project. Do mesmo modo, com os suplementos do Outlook, você pode usar a propriedade [mailbox](http://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) do objeto **Context** para acessar as propriedades e os métodos do objeto **Mailbox** e interagir com a mensagem, a solicitação de reunião ou o conteúdo do compromisso.

O objeto **Context** também fornece acesso às propriedades [contentLanguage](http://dev.office.com/reference/add-ins/shared/office.context.contentlanguage) e [displayLanguage](http://dev.office.com/reference/add-ins/shared/office.context.displaylanguage) que permitem determinar a localidade (idioma) usada no documento ou no item, ou pelo aplicativo host. E a propriedade [roamingSettings](http://dev.office.com/reference/add-ins/outlook/Office.context), que permite acessar os membros do objeto [RoamingSettings](http://dev.office.com/reference/add-ins/outlook/RoamingSettings). Por fim, o objeto **Context** fornece uma propriedade [ui](http://dev.office.com/reference/add-ins/shared/officeui) que permite que o suplemento inicie caixas de diálogo pop-up.


## <a name="document-object"></a>Objeto Documento


 **Aplica-se a:** tipos de suplemento de conteúdo e painel de tarefas

Para interagir com dados do documento no Excel, PowerPoint e Word, a API fornece o objeto [Document](http://dev.office.com/reference/add-ins/shared/document). Você pode usar objetos membros de **Document** para acessar dados das seguintes maneiras:


- Ler e gravar as seleções ativas na forma de texto, células contíguas (matrizes) ou tabelas.
    
- Dados tabulares (matrizes ou tabelas).
    
- Associações (criadas com os métodos "add" do objeto **Bindings**).
    
- Partes XML personalizadas (somente para Word).
    
- Configurações ou estado do suplemento persistido por suplemento no documento.
    
Você também pode usar o objeto **Document** para interagir com os dados nos documentos do Project. A funcionalidade específica do Project para a API está documentada nos membros da classe abstrata [ProjectDocument](http://dev.office.com/reference/add-ins/shared/projectdocument.projectdocument). Para saber mais sobre a criação de suplementos de painel de tarefas, consulte [Suplementos de painel de tarefas para o Project](../project/project-add-ins.md).

Todas essas formas de acesso a dados têm início em uma instância do objeto abstrato **Document**.

Você pode acessar uma instância do objeto **Document** quando o suplemento de painel de tarefas ou de conteúdo for iniciado com a propriedade [document](http://dev.office.com/reference/add-ins/shared/office.context.document) do objeto **Context**. O objeto **Document** define funções comuns do acesso a dados compartilhadas em documentos do Word e do Excel, além de fornecer acesso ao objeto **CustomXmlParts** para documentos do Word.

O objeto **Document** permite que os desenvolvedores acessem o conteúdo de documentos de quatro maneiras:


- Acesso baseado em seleção
    
- Acesso baseado em associação
    
- Acesso baseado em partes personalizadas do XML (apenas para Word)
    
- Acesso baseado em documento (somente para Word e PowerPoint)
    
Para ajudá-lo a entender como os métodos de acesso de dados baseados na seleção e na associação funcionam, explicaremos como as APIs de acesso aos dados proporcionam acesso consistente aos dados de diferentes aplicativos do Office.


### <a name="consistent-data-access-across-office-applications"></a>Acesso consistente aos dados entre aplicativos do Office

 **Aplica-se a:** tipos de suplemento de conteúdo e painel de tarefas

Para criar extensões que funcionam perfeitamente em diferentes documentos do Office, a API JavaScript para Office destaca as particularidades de todos os aplicativos do Office por meio de tipos de dados comuns e da habilidade de forçar diferentes conteúdos de documento para três tipos comuns de dados.


#### <a name="common-data-types"></a>Tipo comuns de dados

Nos acessos a dados baseados em seleção e em associação, os conteúdos dos documentos são expostos por meio dos tipos de dados comuns a todos os aplicativos compatíveis do Office. No Office 2013, há suporte para três tipos de dados principais:



|**Tipo de dados**|**Descrição**|**Suporte ao aplicativo de host**|
|:-----|:-----|:-----|
|Texto|Fornece uma representação, em uma cadeia de caracteres, dos dados na seleção ou associação.|No Excel 2013, no Project 2013 e no PowerPoint 2013, há suporte apenas para texto sem formatação. No Word 2013, há suporte para três formatos de texto: texto sem formatação, HTML e OOXML (Office Open XML). Quando o texto é selecionado em uma célula no Excel, os métodos baseados em seleção realizam os processos de leitura e gravação para todo o conteúdo da célula, mesmo que apenas uma parte do texto esteja selecionada na célula. Quando texto é selecionado no Word e no PowerPoint, os métodos baseados em seleção realizam os processos de leitura e gravação apenas para os caracteres selecionados. O Project 2013 e o PowerPoint 2013 dão suporte apenas ao acesso a dados com base em seleção.|
|Matriz|Fornece os dados na seleção ou associação como uma **Array** bidimensional, que, no JavaScript, é implementada como uma matriz de matrizes. Por exemplo, duas linhas de valores **string** em duas colunas seriam ` [['a', 'b'], ['c', 'd']]`, e uma única coluna com três linhas seria `[['a'], ['b'], ['c']]`.|Há suporte ao acesso a dados de matriz apenas no Excel 2013 e no Word 2013.|
|Tabela|Fornece os dados na seleção ou associação como um objeto [TableData](http://dev.office.com/reference/add-ins/shared/tabledata). O objeto **TableData** expõe os dados por meio de propriedades **headers** e **rows**.|Há suporte ao acesso a dados de tabela apenas no Excel 2013 e no Word 2013.|

#### <a name="data-type-coercion"></a>Coerção de tipo de dados

Os métodos de acesso de dados nos objetos **Document** e [Binding](http://dev.office.com/reference/add-ins/shared/binding) permitem especificar o tipo desejado de dados por meio do parâmetro _coercionType_ desses métodos e os valores de enumeração [CoercionType](http://dev.office.com/reference/add-ins/shared/coerciontype-enumeration) correspondentes. Independentemente da forma real da associação, os diferentes aplicativos do Office dão suporte aos tipos de dados comuns ao tentar forçar os dados a usarem o tipo de dados solicitado. Por exemplo, se uma tabela ou um parágrafo do Word for selecionado, o desenvolvedor pode escolher se deseja lê-lo como texto sem formatação, Office Open XML ou tabela, e a implementação da API manipula as conversões de dados e as transformações necessárias.


 >**Dica**   **Quando devo usar a matriz ou a tabela coercionType para o acesso aos dados?** Se for preciso que os dados tabulares cresçam dinamicamente quando linhas e colunas são adicionadas e você precisar trabalhar com os cabeçalhos da tabela, use o tipo de dados da tabela (especificando o parâmetro _coercionType_ de um método de acesso a dados do objeto **Document** ou **Binding** como `"table"` ou **Office.CoercionType.Table**). A adição de linhas e colunas na estrutura de dados tem suporte nos dados de tabela e matriz, mas o acréscimo de linhas e colunas só tem suporte para dados de tabela. Se você não planeja adicionar linhas e colunas, e os dados não exigem a funcionalidade do cabeçalho, use o tipo de dados de matriz (especificando o parâmetro _coercionType_ do método de acesso a dados como `"matrix"` ou **Office.CoercionType.Matrix**), que fornece um modelo mais simples para interagir com os dados.

Se os dados não puderem ser forçados para o tipo especificado, a propriedade [AsyncResult.status](http://dev.office.com/reference/add-ins/shared/asyncresult.error) presente nos retornos de chamada retorna `"failed"`, e você pode usar a propriedade [AsyncResult.error](http://dev.office.com/reference/add-ins/shared/asyncresult.context) para acessar um objeto [Error](http://dev.office.com/reference/add-ins/shared/error) com informações sobre o motivo pelo qual a chamada de método falhou.


## <a name="working-with-selections-using-the-document-object"></a>Trabalhar com seleções que usam o objeto Document


O objeto **Document** expõe métodos que permitem ler e gravar a seleção atual do usuário de uma maneira "obter e definir". Para fazer isso, o objeto **Document** fornece os métodos **getSelectedDataAsync** e **setSelectedDataAsync**.

Para obter exemplos de códigos que demostram como realizar tarefas com seleções, consulte [Ler e gravar dados na seleção ativa em um documento ou uma planilha](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).


## <a name="working-with-bindings-using-the-bindings-and-binding-objects"></a>Trabalhar com associações usando os objetos Bindings e Binding


O acesso a dados baseado em associação habilita os suplementos de conteúdo e painel de tarefas a acessarem de forma consistente determinada região de um documento ou uma planilha por meio de um identificador vinculado a uma associação. Primeiro, o suplemento precisa estabelecer a associação chamando um dos métodos que vinculam uma parte do documento a um identificador exclusivo: [addFromPromptAsync](http://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync), [addFromSelectionAsync](http://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync) ou [addFromNamedItemAsync](http://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync). Depois que a associação é estabelecida, o suplemento pode usar o identificador fornecido para acessar os dados contidos na região vinculada do documento ou da planilha. A criação de associações fornece o seguinte valor ao suplemento:


- Permite o acesso a estruturas comuns de dados em aplicativos compatíveis do Office, como: tabelas, intervalos ou texto (uma execução contígua de caracteres).
    
- Habilita operações de leitura/gravação sem exigir que o usuário realize uma seleção.
    
- Estabelece uma relação entre o suplemento e os dados presentes no documento. As associações estão presentes no documento e podem ser acessadas em um momento posterior.
    
A criação de uma associação também permite que você se inscreva em eventos de alteração de seleção e de dados que apresentem um escopo definido para essa região específica do documento ou da planilha. Isso significa que o suplemento só é notificado sobre alterações que ocorrem dentro da região associada, e não sobre alterações gerais que ocorrem em todo o documento ou planilha.

O objeto [Bindings](http://dev.office.com/reference/add-ins/shared/bindings.bindings) expõe um método [getAllAsync](http://dev.office.com/reference/add-ins/shared/bindings.getallasync), que dá acesso ao conjunto de todas as associações estabelecidas no documento ou planilha. Uma associação individual pode ser acessada por sua ID. Para isso, use os métodos [Bindings.getBindingByIdAsync](http://dev.office.com/reference/add-ins/shared/bindings.getbyidasync) ou [Office.select](http://dev.office.com/reference/add-ins/shared/office.select). Você pode estabelecer novas associações e remover as associações existentes usando um dos seguintes métodos para o objeto **Bindings**: [addFromSelectionAsync](http://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync), [addFromPromptAsync](http://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync), [addFromNamedItemAsync](http://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync) ou [releaseByIdAsync](http://dev.office.com/reference/add-ins/shared/bindings.releasebyidasync).

Há três tipos diferentes de associações que podem ser especificadas com o parâmetro _bindingType_ durante a criação de uma nova associação com os métodos **addFromSelectionAsync**, **addFromPromptAsync** ou **addFromNamedItemAsync**:



|**Tipo de associação**|**Descrição**|**Suporte ao aplicativo de host**|
|:-----|:-----|:-----|
|Associação de texto|Associa a uma região do documento que pode ser representada como um texto.|No Word, a maioria das seleções contíguas são válidas, enquanto no Excel apenas as seleções de células únicas podem ser usadas para uma associação de texto. No Excel, só há suporte para texto sem formatação. No Word, há suporte para três formatos: texto sem formatação, HTML e Open XML do Office.|
|Associação de matriz|Associa a uma região fixa de um documento que contém dados tabulares sem cabeçalhos. Os dados de uma associação de matriz são gravados ou lidos como uma **Array** bidimensional, que é implementada como uma matriz de matrizes no JavaScript. Por exemplo, duas linhas de valores **string** em duas colunas podem ser gravadas ou lidas como ` [['a', 'b'], ['c', 'd']]`, e uma única coluna de três linhas pode ser gravada ou lida como `[['a'], ['b'], ['c']]`.|No Excel, qualquer seleção contígua de células pode ser usada para estabelecer uma associação de matriz. No Word, apenas as tabelas dão suporte à associação de matriz.|
|Associação de tabelas|Associa a uma região de um documento que contém uma tabela com cabeçalhos. Os dados em uma associação de tabela são gravados ou lidos como um objeto [TableData](http://dev.office.com/reference/add-ins/shared/tabledata). O objeto **TableData** expõe os dados por meio das propriedades **headers** e **rows**.|Qualquer tabela do Excel ou Word pode ser a base para uma associação de tabela. Após estabelecer uma associação de tabelas, as linhas ou colunas novas que um usuário adicionar à tabela são automaticamente incluídas na associação.  |
Depois que uma associação é criada usando um dos três métodos "add" do objeto **Bindings**, é possível trabalhar com os dados e as propriedades da associação usando os métodos do objeto correspondente: [MatrixBinding](http://dev.office.com/reference/add-ins/shared/binding.matrixbinding), [TableBinding](http://dev.office.com/reference/add-ins/shared/binding.tablebinding) ou [TextBinding](http://dev.office.com/reference/add-ins/shared/binding.textbinding). Esses três objetos herdam os métodos [getDataAsync](http://dev.office.com/reference/add-ins/shared/binding.getdataasync) e [setDataAsync](http://dev.office.com/reference/add-ins/shared/binding.setdataasync) do objeto **Binding**, o que o habilita a interagir com os dados associados.

Para obter exemplos de códigos que demonstram como realizar tarefas com associações, consulte [Associar a regiões em um documento ou uma planilha](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md).


## <a name="working-with-custom-xml-parts-using-the-customxmlparts-and-customxmlpart-objects"></a>Trabalhar com partes XML personalizadas usando os objetos CustomXmlParts e CustomXmlPart


 **Aplica-se a:** suplementos de painel de tarefas para Word

Os objetos [CustomXmlParts](http://dev.office.com/reference/add-ins/shared/customxmlparts.customxmlparts) e [CustomXmlPart](http://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart) da API fornecem acesso a partes XML personalizadas em documentos do Word, que habilitam a manipulação orientada por XML do conteúdo do documento. Para obter demonstrações de como trabalhar com objetos **CustomXmlParts** e **CustomXmlPart**, consulte o exemplo de código [Word-Add-in-Work-with-custom-XML-parts](https://github.com/OfficeDev/Word-Add-in-Work-with-custom-XML-parts).


## <a name="working-with-the-entire-document-using-the-getfileasync-method"></a>Trabalhar com o documento inteiro usando o método getFileAsync


 **Aplica-se a:** suplementos de painel de tarefas para Word e PowerPoint

O método [Document.getFileAsync](http://dev.office.com/reference/add-ins/shared/document.getfileasync) e os membros dos objetos [File](http://dev.office.com/reference/add-ins/shared/file) e [Slice](http://dev.office.com/reference/add-ins/shared/slice) fornecem a funcionalidade necessária para obter documentos inteiros do Word e PowerPoint em fatias (frações) de até 4 MB por vez. Para saber mais, consulte [Como obter todo o conteúdo do arquivo a partir de um documento em um suplemento](../../docs/develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word.md).


## <a name="mailbox-object"></a>Objeto Mailbox


 **Aplica-se a:** suplementos do Outlook

Os suplementos do Outlook usam principalmente um subconjunto da API exposta no objeto [Mailbox](http://dev.office.com/reference/add-ins/outlook/Office.context.mailbox). Para acessar os objetos e membros específicos para suplementos do Outlook, como o objeto [Item](http://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item), use a propriedade [mailbox](http://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) do objeto **Context** para acessar o objeto **Mailbox**, conforme exibido na linha de código abaixo.




```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

Além disso, os suplementos do Outlook podem usar os seguintes objetos:


-  Objeto **Office**: para inicialização.
    
-  Objeto **Context**: para acesso a propriedades de conteúdo e idioma de exibição.
    
-  Objeto **RoamingSettings**: para salvar as configurações personalizadas do suplemento do Outlook na caixa de correio do usuário em que o suplemento está instalado.
    
Para saber mais sobre como usar o JavaScript em suplementos do Outlook, consulte [Suplementos do Outlook](../outlook/outlook-add-ins.md) e [Visão geral da arquitetura e dos recursos de suplementos do Outlook](../outlook/overview.md).


## <a name="api-support-matrix"></a>Matriz de suporte da API


Esta tabela resume a API e os recursos compatíveis com os tipos de suplemento (conteúdo, painel de tarefas e Outlook) e os aplicativos do Office que podem hospedá-los quando o usuário especifica os [aplicativos hospedados pelo Office compatíveis com o suplemento](http://msdn.microsoft.com/library/cff9fbdf-a530-4f6e-91ca-81bcacd90dcd%28Office.15%29.aspx) usando o [esquema 1.1 do manifesto de suplementos e recursos compatíveis com a v1.1 da API JavaScript para Office](../../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md).


|||||||||
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
||**Nome do host**|Banco de dados|Pasta de trabalho|Caixa de correio|Apresentação|Documento|Project|
||**Aplicativos host** **compatíveis**|Aplicativos Web do Access|ExcelExcel Online|OutlookOutlook Web AppOWA para Dispositivos|PowerPointPowerPoint Online|Word|Project|
|**Tipos de suplemento com suporte**|Conteúdo|S|S||S|||
||Painel de tarefas||S||S|S|S|
||Outlook|||S||||
|**Recursos da API compatíveis**|Ler/gravar texto||S||S|S|S (somente leitura)|
||Ler/gravar matriz||S|||S||
||Ler/gravar tabela||S|||S||
||Ler/gravar HTML|||||S||
||Ler/gravar Office Open XML|||||S||
||Ler propriedades de tarefa, recurso, modo de exibição e campo||||||S|
||Eventos alterados pela seleção||S|||S||
||Obter documento inteiro||||S|S||
||Associações e eventos de associação|S (somente associações de tabela totais e parciais)|S|||S||
||Ler/gravar partes XML personalizadas|||||S||
||Persistir dados de estado de suplemento (configurações)|S (por suplemento de host)|S (por documento)|S (por caixa de correio)|S (por documento)|S (por documento)||
||Eventos alterados pelas configurações|S|S||S|S||
||Obter o modo de exibição ativo e visualizar eventos alterados||||S|||
||Navegar para locais no documento||S||S|S||
||Ativar contextualmente usando regras e RegEx|||S||||
||Ler propriedades do item|||S||||
||Ler perfil de usuário|||S||||
||Obter anexos|||S||||
||Obter o token de identidade do usuário|||S||||
||Chamar os serviços Web do Exchange|||S||||
