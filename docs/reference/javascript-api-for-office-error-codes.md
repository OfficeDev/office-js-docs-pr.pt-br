# <a name="javascript-api-for-office-error-codes"></a>Códigos de erro da API do JavaScript para Office

Este artigo documenta as mensagens de erro que você pode receber ao usar a API JavaScript para Office (Office.js).

**Aplica-se a:** Suplementos do Office | Suplementos do SharePoint | Excel | Outlook | PowerPoint | Project | Word

## <a name="error-codes"></a>Códigos de erro

A tabela a seguir lista os códigos de erro, nomes e mensagens exibidas e as condições que indicam.

|**Error.code**|**Error.name**|**Error.message**|**Condição**|
|:-----|:-----|:-----|:-----|
|1000|Tipo de coerção inválido|O tipo de coerção especificado não tem suporte|O tipo de coerção não é suportado no aplicativo host. (Por exemplo, os tipos de coerção OOXML e HTML não são suportados no Excel.)|
|1001|Erro de leitura de dados|A seleção atual não tem suporte.|A seleção atual do usuário não é suportada. (Isto é, é algo diferente dos tipos de coerção suportados.)|
|1002|Tipo de coerção inválido|O tipo de coerção especificado não é compatível para esse tipo de ligação.|O desenvolvedor da solução forneceu uma combinação incompatível de tipo de coerção e tipo de associação.|
|1003|Erro de leitura de dados|Os valores rowCount ou columnCount especificados são inválidos.|O usuário fornece contagens inválidas de coluna ou de linha.|
|1004|Erro de leitura de dados|A seleção atual não tem suporte para o tipo de coerção especificado.|A seleção atual não é compatível com o tipo de coerção especificado por este aplicativo.|
|1005|Erro de leitura de dados|Os valores startRow ou startColumn especificados são inválidos.|O usuário fornece valores inválidos de startRow ou startCol.|
|1006|Erro de leitura de dados|Os parâmetros de coordenadas não podem ser usados com o tipo de coerção “Table” quando a tabela contiver células mescladas.|O usuário tenta obter dados parciais de uma tabela não uniforme (ou seja, uma tabela que tenha células mescladas). |
|1007|Erro de leitura de dados|O tamanho do documento é muito grande.|O usuário tentar obter um documento maior do que o tamanho compatível no momento.|
|1008|Erro de leitura de dados|O conjunto de dados solicitado é muito grande.|O usuário solicita a leitura de dados além dos limites de dados definidos pelos suplementos host.|
|1009|Erro de leitura de dados|O tipo de arquivo especificado não tem suporte.|O usuário envia um tipo de arquivo inválido.|
|2000|Erro de gravação de dados|Não há suporte para o tipo de objeto de dados fornecido. |Um objeto de dados sem suporte foi fornecido.|
|2001|Erro de gravação de dados|Não é possível gravar na seleção atual.|Não há suporte para a seleção atual do usuário para uma operação de gravação. (Por exemplo, quando o usuário seleciona uma imagem.)|
|2002|Erro de gravação de dados|O objeto de dados fornecido não é compatível com a forma ou com as dimensões da seleção atual.|Várias células estão selecionadas (e a forma de seleção não coincide com a forma dos dados). Várias células estão selecionadas (e as dimensões da seleção não coincidem com as dimensões dos dados).|
|2003|Erro de gravação de dados|A operação de configuração falhou porque o objeto de dados fornecido substituirá os dados.|Uma única célula está selecionada e o objeto de dados fornecido substitui os dados na planilha.|
|2004|Erro de gravação de dados|O objeto de dados fornecido não corresponde ao tamanho da seleção atual.|O usuário fornece um objeto maior do que o tamanho da seleção atual.|
|2005|Erro de gravação de dados|Os valores startRow ou startColumn especificados são inválidos.|O usuário fornece valores inválidos de startRow ou startCol.|
|2006|Erro de formato inválido|O formato do objeto de dados especificado não é válido.|O desenvolvedor da solução fornece uma cadeia HTML ou OOXML inválida, uma cadeia HTML malformada ou uma cadeia OOXML inválida.|
|2007|Objeto de dados inválido|O tipo do objeto de dados especificado não é compatível com a seleção atual.|O desenvolvedor da solução fornece um objeto de dados incompatível com o tipo de coerção especificado.|
|2008|Erro de gravação de dados|TBD|TBD|
|2009|Erro de gravação de dados|O objeto de dados especificado é muito grande.|O usuário tenta definir dados além dos limites de dados definidos pelos suplementos do host.|
|2010|Erro de gravação de dados|Os parâmetros de coordenadas não podem ser usados com o tipo de coerção Table quando a tabela contiver células mescladas.|O usuário tentar definir dados parciais de uma tabela não uniforme (ou seja, uma tabela que possui células mescladas).|
|3000|Erro de criação de associação|Não é possível associar à seleção atual.|A seleção do usuário não é suportada para associação. (Por exemplo, o usuário está selecionando uma imagem ou outro objeto sem suporte.)|
|3001|Erro de criação de associação|TBD|TBD|
|3002|Erro de associação inválida|A associação especificada não existe.|O desenvolvedor tenta associar a uma associação não existente ou removida.|
|3003|Erro de criação de associação|Não há suporte para várias seleções não contíguas.|O usuário está fazendo várias seleções.|
|3004|Erro de criação de associação|Não é possível criar uma associação com a seleção atual e o tipo de associação especificada.|Há várias condições em que isso pode acontecer. Confira a seção "Condições de erro de criação de associação" posteriormente neste artigo.|
|3005|Operação de associação inválida|Operação sem suporte neste tipo de associação.|O desenvolvedor envia uma operação de adição de linha ou de adição de coluna em um tipo de associação que não é _table_.|
|3006|Erro de criação de associação|O item nomeado não existe.|Não foi possível encontrar o item nomeado. Não existe um controle de conteúdo ou uma tabela com esse nome.|
|3007|Erro de criação de associação|Foram encontrados vários objetos com o mesmo nome.|Erro de colisão: há mais de um controle de conteúdo com o mesmo nome e a falha na colisão está definida como **true**.|
|3008|Erro de criação de associação|O tipo de associação especificado não é compatível com o item nomeado fornecido.|Não é possível associar o item nomeado ao tipo. Por exemplo, um controle de conteúdo contém texto, mas o desenvolvedor tentou associar usando o tipo de coerção _table_.|
|3009|Operação de associação inválida|Não há suporte para o tipo de associação.|Usado para fins de compatibilidade com versões anteriores.|
|3010|Operação de associação inválida|O conteúdo selecionado precisa estar em formato de tabela. Formate os dados como uma tabela e tente novamente.|O desenvolvedor está tentando usar os métodos **addRowsAsynch** ou **deleteAllDataValuesAsynch** do objeto **TableBinding** em dados do tipo de coerção _matrix_.|
|4000|Erro de leitura de configurações|O nome de configuração especificado não existe.|Um nome de configuração inexistente foi fornecido.|
|4001|Erro ao salvar configurações|Não foi possível salvar as configurações.|Não foi possível salvar as configurações.|
|4002|Erro de configurações obsoletas|Não foi possível salvar as configurações porque elas estão obsoletas.|As configurações estão obsoletas e o desenvolvedor indicou que não devem ser substituídas.|
|5000|Erro de configurações obsoletas|Não há suporte para a operação.|A operação não tem suporte no host atual. Por exemplo, **document.getSelectionAsync** é chamada do Outlook.|
|5001|Erro interno|Ocorreu um erro interno.|Refere-se a uma condição de erro interno, que pode ocorrer por qualquer um dos seguintes motivos:<br/><table><tr><td>Um suplemento que está sendo usado por outro usuário que compartilha a pasta de trabalho criada com uma associação aproximadamente no mesmo momento, e seu suplemento precisa tentar realizar a associação novamente.</tr></td><tr><td>Ocorreu um erro desconhecido.</tr></td><tr><td>Falha na operação.</tr></td><tr><td>O acesso foi negado porque o usuário não é um membro de uma função autorizada.</tr></td><tr><td>O acesso foi negado porque é necessária a comunicação segura e criptografada.</tr></td><tr><td>Os dados estão obsoletos e o usuário precisa confirmar permitindo que as consultas os atualizem.</tr></td><tr><td>A cota de CPU do conjunto de sites foi excedida.</tr></td><tr><td>A cota de memória do conjunto de sites foi excedida.</tr></td><tr><td>A cota de memória da sessão foi excedida.</tr></td><tr><td>A pasta de trabalho está em um estado inválido e a operação não pode ser executada.</tr></td><tr><td>A sessão expirou devido a inatividade e o usuário precisa recarregar a pasta de trabalho.</tr></td><tr><td>A quantidade máxima de sessões permitida por usuário foi excedida.</tr></td><tr><td>A operação foi cancelada pelo usuário.</tr></td><tr><td>Não foi possível concluir a operação porque ela está demorando muito.</tr></td><tr><td>Não foi possível concluir a solicitação e ela deve ser repetida.</tr></td><tr><td>O período de avaliação do produto expirou.</tr></td><tr><td>A sessão expirou devido a inatividade.</tr></td><tr><td>O usuário não tem permissão para executar a operação no intervalo especificado.</tr></td><tr><td>As configurações regionais do usuário não correspondem às da sessão atual de colaboração.</tr></td><tr><td>O usuário não está mais conectado e deve atualizar ou abrir novamente a pasta de trabalho.</tr></td><tr><td>O intervalo solicitado não existe na planilha.</tr></td><tr><td>O usuário não tem permissão para editar a pasta de trabalho.</tr></td><tr><td>Não foi possível editar a pasta de trabalho porque ela está bloqueada.</tr></td><tr><td>A sessão não pode salvar a pasta de trabalho automaticamente.</tr></td><tr><td>A sessão não pode atualizar seu bloqueio no arquivo de pasta de trabalho.</tr></td><tr><td>Não foi possível processar a solicitação e ela deve ser repetida.</tr></td><tr><td>Não foi possível verificar as informações de entrada do usuário e elas precisam ser inseridas novamente.</tr></td><tr><td>O usuário teve o acesso negado.</tr></td><tr><td>A pasta de trabalho compartilhada precisa ser atualizada.</tr></td></table>|
|5002|Permissão negada|A operação solicitada não é permitida no modo de documento atual.|O desenvolvedor da solução envia uma operação de definição, mas o documento está em um modo que não permite alterações, como “Restringir Edição”.|
|5003|Erro de registro de eventos|Não há suporte para o tipo de evento especificado pelo objeto atual.|O desenvolvedor da solução tenta registrar ou cancelar o registro de um manipulador para um evento que não existe.|
|5004|Chamada de API inválida|Chamada à API inválida no contexto atual.|Uma chamada inválida foi feita para o contexto, por exemplo, tentando usar um objeto **CustomXMLPart** no Excel.|
|5005|Dados obsoletos|Falha na operação devido aos dados estarem obsoletos no servidor.|Os dados no servidor precisam ser atualizados.|
|5006|Tempo limite da sessão|O tempo limite da sessão do documento esgotou-se. Recarregue o documento. |A sessão expirou.|
|5007|Chamada de API inválida|Não há suporte à enumeração no contexto atual.|Não há suporte à enumeração no contexto atual.|
|5009|Permissão negada|Acesso negado|O suplemento não tem permissão para chamar a API específica.|
|6000|Nó inválido|O nó especificado não foi encontrado.|O nó **CustomXmlPart** não foi encontrado.|
|6100|Erro no XML personalizado|Erro no XML personalizado|Chamada de API inválida|
|7000|ID inválida|A ID especificada não existe.|ID inválida|
|7001|Navegação inválida|O objeto está localizado em um local em onde a navegação não é suportada.|O usuário pode encontrar o objeto, mas não é possível navegar até ele. (Por exemplo, no Word, a associação ocorre com o cabeçalho, rodapé ou um comentário.)|
|7002|Navegação inválida|O objeto está bloqueado ou protegido.|O usuário está tentando navegar até um intervalo bloqueado ou protegido.|
|7004|Navegação inválida|A operação falhou porque o índice está fora do intervalo.|O usuário está tentando navegar até um índice fora do intervalo.|
|8000|Parâmetro ausente|Não foi possível formatar a célula da tabela porque faltam alguns valores de parâmetro. Verifique os parâmetros e tente novamente.|Faltam alguns parâmetros no método cellFormat. Por exemplo, faltam células, formatações ou parâmetros de tableOptions.|
|8010|Valor inválido|Um ou mais parâmetros de célula têm valores que não são permitidos. Verifique os valores e tente novamente.|A enumeração de referência de células comuns não está definida. Por exemplo, Todos, Dados, Cabeçalhos.|
|8011|Valor inválido|Um ou mais parâmetros tableOptions possuem valores que não são permitidos. Verifique os valores e tente novamente.|Um dos valores em tableOptions é inválido.|
|8012|Valor inválido|Um ou mais parâmetros de formato possuem valores que não são permitidos. Verifique os valores e tente novamente.|Um dos valores no formato é inválido.|
|8020|Fora do intervalo|O valor de índice de linha está fora do intervalo permitido. Use um valor positivo (0 ou maior) que seja menor do que o número de linhas.|O índice de linha é superior ao maior índice de linha da tabela ou menor do que 0.|
|8021|Fora do intervalo|O valor de índice de coluna está fora do intervalo permitido. Use um valor positivo (0 ou maior) que seja menor do que o número de colunas.|O índice de coluna é superior ao maior índice de coluna da tabela ou menor do que 0.|
|8022|Fora do intervalo|O valor está fora do intervalo permitido.|Alguns dos valores no formato estão fora dos intervalos suportados.|
|9016|Permissão negada|Permissão negada|Acesso negado.|
|12002|||Uma das seguintes opções:<br> - Não existe uma página na URL transmitida para `displayDialogAsync`.<br> - A página transmitida para `displayDialogAsync` foi carregada, mas a caixa de diálogo foi direcionada para uma página que ela não consegue localizar nem carregar ou foi direcionada para uma URL com sintaxe inválida. Gerado na caixa de diálogo e dispara um evento `DialogEventReceived` na página host.|
|12003|||A caixa de diálogo foi direcionada para uma URL com o protocolo HTTP. HTTPS é obrigatório. Gerado na caixa de diálogo e dispara um evento `DialogEventReceived` na página host.|
|12004|||O domínio da URL é passado para `displayDialogAsync` não é confiável. O domínio deve ser do mesmo domínio que a página de host (incluindo o número da porta e protocolo). Gerado por chamada de `displayDialogAsync`.|
|12005|||A URL passada para `displayDialogAsync` usa o protocolo HTTP. HTTPS é obrigatório. Gerado por chamada de `displayDialogAsync`. (Em algumas versões do Office, a mensagem de erro retornada para 12005 é a mesma retornada para 12004.)|
|12006|||A caixa de diálogo foi fechada, geralmente pelo usuário ter escolhido o botão **X**. Gerado na caixa de diálogo e dispara um evento `DialogEventReceived` na página host.|
|12007|||Uma caixa de diálogo já está aberta a partir desta janela do host. Uma janela de host, como um painel de tarefas, só pode ter uma caixa de diálogo aberta por vez. Gerado por chamada de `displayDialogAsync`.|
|13000 - 13010|||Comfira [Causas e tratamento dos erros do getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/troubleshoot-sso-in-office-add-ins#causes-and-handling-of-errors-from-getaccesstokenasync).|

## <a name="binding-creation-error-conditions"></a>Condições do erro de criação de associação

Quando uma associação for criada na API, indique o tipo de associação que deseja usar. A tabela a seguir lista os tipos de associação e os comportamentos de associação resultantes esperados.

### <a name="behavior-in-excel"></a>Comportamento no Excel

A tabela a seguir resume o comportamento de associação no Excel.

|**Tipo de associação especificada**|**Seleção real**|**Comportamento**|
|:-----|:-----|:-----|
|Matrix|Intervalo de células (incluindo dentro de uma tabela e uma única célula)|Uma associação de tipo _matrix_ é criada nas células selecionadas. Nenhuma modificação no documento é esperada.|
|Matrix|Texto selecionado na célula|Uma associação de tipo _matrix_ é criada na célula inteira. Nenhuma modificação no documento é esperada.|
|Matrix|Várias seleções/seleção inválida (por exemplo, o usuário seleciona uma imagem, objeto ou WordArt.)|Não é possível criar a associação.|
|Table|Intervalo de células (inclui uma única célula)|Não é possível criar a associação.|
|Table|Intervalo de células dentro de uma tabela (inclui uma única célula dentro de uma tabela, ou a tabela inteira ou texto dentro de uma célula em uma tabela)|Uma associação é criada na tabela inteira.|
|Table|Metade da seleção em uma tabela e metade da seleção fora da tabela|Não é possível criar a associação.|
|Table|Texto selecionado na célula (não na tabela).|Não é possível criar a associação.|
|Table|Várias seleções/seleção inválida (por exemplo, o usuário seleciona uma imagem, objeto, WordArt, etc.)|Não é possível criar a associação.|
|Text|Intervalo de células|Não é possível criar a associação.|
|Text|Intervalo de células dentro de uma tabela|Não é possível criar a associação.|
|Text|Célula única|Uma associação do tipo _text_ é criada.|
|Text|Célula única dentro de uma tabela|Uma associação do tipo _text_ é criada.|
|Text|Texto selecionado na célula|Uma associação do tipo _text_ é criada na célula inteira.|

### <a name="behavior-in-word"></a>Comportamento no Word

A tabela a seguir resume o comportamento de associação no Word.

|**Tipo de associação especificada**|**Seleção real**|**Comportamento**|
|:-----|:-----|:-----|
|Matrix|Text|Não é possível criar a associação.|
|Matrix|Tabela inteira|Uma associação do tipo _matrix_ é criada. O documento é alterado e um controle de conteúdo deve encapsular a tabela. |
|Matrix|Intervalo dentro de uma tabela|Não é possível criar a associação.|
|Matrix|Seleção inválida (por exemplo, múltiplos objetos, objetos inválidos, etc.)|Não é possível criar a associação.|
|Table|Text|Não é possível criar a associação.|
|Table|Tabela inteira|Uma associação do tipo _text_ é criada.|
|Table|Intervalo dentro de uma tabela|Não é possível criar a associação.|
|Table|Seleção inválida (por exemplo, múltiplos objetos, objetos inválidos, etc.)|Não é possível criar a associação.|
|Text|Tabela inteira|Uma associação do tipo _text_ é criada.|
|Text|Intervalo dentro de uma tabela|Não é possível criar a associação.|
|Text|Seleção múltipla|A última seleção será encapsulada com um controle de conteúdo e uma associação a esse controle. Um controle de conteúdo do tipo _text_ é criado.|
|Text|Seleção inválida (por exemplo, múltiplos objetos, objetos inválidos, etc.)|Não é possível criar a associação.|

## <a name="see-also"></a>Confira também
   
- [Ciclo de vida de desenvolvimento de suplementos do Office](https://docs.microsoft.com/office/dev/add-ins/concepts/add-in-development-lifecycle)
    
