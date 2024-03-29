---
title: Códigos de erro comuns da API do Office
description: Este artigo documenta as mensagens de erro que você pode encontrar ao usar a API Comum do Office.
ms.date: 09/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: d77b4c0c458e11da0057f06a5088ef8a28e4ccd2
ms.sourcegitcommit: 05be1086deb2527c6c6ff3eafcef9d7ed90922ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/28/2022
ms.locfileid: "68092977"
---
# <a name="office-common-api-error-codes"></a>Códigos de erro comuns da API do Office

Este artigo documenta as mensagens de erro que você pode encontrar ao usar o modelo de API Comum. Esses códigos de erro não se aplicam a APIs específicas do aplicativo, como a API JavaScript do Excel ou a API JavaScript do Word.

Consulte [os modelos de API](../develop/understanding-the-javascript-api-for-office.md#api-models) para saber mais sobre as diferenças entre a API Comum e os modelos de API específicos do aplicativo.

## <a name="error-codes"></a>Códigos de erro

A tabela a seguir lista os códigos de erro, nomes e mensagens exibidas e as condições que indicam.

|Error.code|Error.name|Error.message|Condição|
|:-----|:-----|:-----|:-----|
|1000|Tipo inválido de coerção|O tipo de coerção especificado não tem suporte|Não há suporte para o tipo de coerção no aplicativo do Office. (Por exemplo, não há suporte para os tipos de coerção OOXML e HTML no Excel.)|
|1001|Erro de Leitura de Dados|A seleção atual não tem suporte.|Não há suporte para a seleção atual do usuário (ou seja, é algo diferente dos tipos de coerção com suporte).|
|1002|Tipo inválido de coerção|O tipo de coerção especificado não é compatível com este tipo de associação.|O desenvolvedor da solução forneceu uma combinação incompatível de tipo de coerção e tipo de associação.|
|1003|Erro de Leitura de Dados|Os valores rowCount ou columnCount especificados são inválidos.|O usuário fornece contagens inválidas de coluna ou de linha.|
|1004|Erro de Leitura de Dados|A seleção atual não tem suporte para o tipo de coerção especificado.|A seleção atual não é compatível com o tipo de coerção especificado por este aplicativo.|
|1005|Erro de Leitura de Dados|Os valores startRow ou startColumn especificados são inválidos.|O usuário fornece valores inválidos de startRow ou startCol.|
|1006|Erro de Leitura de Dados|Os parâmetros de coordenadas não podem ser usados com o tipo de coerção “Table” quando a tabela contiver células mescladas.|O usuário tenta obter dados parciais de uma tabela não uniforme (ou seja, uma tabela que possui células mescladas). |
|1007|Erro de Leitura de Dados|O tamanho do documento é muito grande.|O usuário tentar obter um documento maior do que o tamanho compatível no momento.|
|1008|Erro de Leitura de Dados|O conjunto de dados solicitado é muito grande.|O usuário solicita a leitura de dados além dos limites de dados definidos pelo aplicativo do Office.|
|1009|Erro de Leitura de Dados|O tipo de arquivo especificado não tem suporte.|O usuário envia um tipo de arquivo inválido.|
|2000|Erro de Gravação de Dados|Não há suporte para o tipo de objeto de dados fornecido. |Um objeto de dados sem suporte foi fornecido.|
|2001|Erro de Gravação de Dados|Não é possível gravar na seleção atual.|The user's current selection is not supported for a write operation. (For example, when the user selects an image.)|
|2002|Erro de Gravação de Dados|O objeto de dados fornecido não é compatível com a forma ou com as dimensões da seleção atual.|Várias células são selecionadas (e a forma de seleção não corresponde à forma dos dados). Várias células são selecionadas (e as dimensões da seleção não correspondem às dimensões dos dados).|
|2003|Erro de Gravação de Dados|A operação de definição falhou porque o objeto de dados fornecido substituirá os dados.|Uma única célula está selecionada e o objeto de dados fornecido substitui os dados na planilha.|
|2004|Erro de Gravação de Dados|O objeto de dados fornecido não corresponde ao tamanho da seleção atual.|O usuário fornece um objeto maior do que o tamanho da seleção atual.|
|2005|Erro de Gravação de Dados|Os valores startRow ou startColumn especificados são inválidos.|O usuário fornece valores inválidos de startRow ou startCol.|
|2006|Erro de formato inválido|O formato do objeto de dados especificado não é válido.|O desenvolvedor de solução fornece uma cadeia de caracteres HTML ou OOXML inválida, uma cadeia de caracteres HTML mal formada ou uma cadeia de caracteres OOXML inválida.|
|2007|Objeto de dados inválido|O tipo do objeto de dados especificado não é compatível com a seleção atual.|O desenvolvedor da solução fornece um objeto de dados incompatível com o tipo de coerção especificado.|
|2008|Erro de Gravação de Dados|TBD|TBD|
|2009|Erro de Gravação de Dados|O objeto de dados especificado é muito grande.|O usuário tenta definir dados além dos limites de dados definidos pelo aplicativo do Office.|
|2010|Erro de Gravação de Dados|Os parâmetros de coordenadas não podem ser usados com o tipo de coerção Table quando a tabela contiver células mescladas.|O usuário tenta definir dados parciais de uma tabela não uniforme (ou seja, uma tabela que possui células mescladas).|
|3000|Erro de Criação de Associação|Não é possível associar à seleção atual.|The user's selection is not supported for binding. (For example, the user is selecting an image or other non-supported object.)|
|3001|Erro de Criação de Associação|TBD|TBD|
|3002|Erro de Associação Inválida|A ligação especificada não existe.|O desenvolvedor tenta associar a uma associação não existente ou removida.|
|3003|Erro de Criação de Associação|Não há suporte para várias seleções não contíguas.|O usuário está fazendo várias seleções.|
|3004|Erro de Criação de Associação|Não é possível criar uma ligação com a seleção atual e o tipo de ligação especificada.|There are several conditions under which this might happen. Please see the "Binding creation error conditions" section later in this article.|
|3005|Operação de Associação Inválida|Operação sem suporte neste tipo de associação.|O desenvolvedor envia uma operação adicionar linha ou adicionar coluna em um tipo de associação que não é do tipo de coerção `table`.|
|3006|Erro de Criação de Associação|O item nomeado não existe.|The named item cannot be found. No content control or table with that name exists.|
|3007|Erro de Criação de Associação|Foram encontrados vários objetos com o mesmo nome.|Erro de colisão: existe mais de um controle de conteúdo com o mesmo nome e a falha na colisão é definida como `true`.|
|3008|Erro de Criação de Associação|O tipo de associação especificado não é compatível com o item nomeado fornecido.|O item nomeado não pode ser associado ao tipo. Por exemplo, um controle de conteúdo contém texto, mas o desenvolvedor tentou associar usando o tipo de coerção `table`.|
|3009|Operação de Associação Inválida|Não há suporte para o tipo de vinculação.|Usado para fins de compatibilidade com versões anteriores.|
|3010|Operação de Associação Inválida|O conteúdo selecionado precisa estar em formato de tabela. Formate os dados como uma tabela e tente novamente.|O desenvolvedor está tentando usar o método `addRowsAsync` ou o `deleteAllDataValuesAsync` objeto em `TableBinding` dados do tipo de coerção `matrix`.|
|4000|Erro de leitura de configurações|O nome de configuração especificado não existe.|Um nome de configuração inexistente foi fornecido.|
|4001|Salvar erro de configurações|Não foi possível salvar as configurações.|Não foi possível salvar as configurações.|
|4002|Erro de configurações obsoletos|Não foi possível salvar as configurações porque elas estão obsoletas.|As configurações estão obsoletas e o desenvolvedor indicou que não devem ser substituídas.|
|5000|Erro de configurações obsoletos|Não há suporte para a operação.|Não há suporte para a operação no aplicativo atual do Office. Por exemplo, é `document.getSelectionAsync` chamado do Outlook.|
|5001|Erro interno|Ocorreu um erro interno.|Refere-se a uma condição de erro interna, que pode ocorrer por qualquer um dos motivos a seguir.<br/><table><tr><td>Um suplemento que está sendo usado por outro usuário que compartilha a pasta de trabalho criada com uma associação aproximadamente no mesmo momento, e seu suplemento precisa tentar realizar a associação novamente.</tr></td><tr><td>Ocorreu um erro desconhecido.</tr></td><tr><td>Falha na operação.</tr></td><tr><td>O acesso foi negado porque o usuário não é um membro de uma função autorizada.</tr></td><tr><td>O acesso foi negado porque é necessária a comunicação segura e criptografada.</tr></td><tr><td>Os dados estão obsoletos, e o usuário precisa confirmar permitindo que as consultas os atualizem.</tr></td><tr><td>A cota de CPU do conjunto de sites foi excedida.</tr></td><tr><td>A cota de memória do conjunto de sites foi excedida.</tr></td><tr><td>A cota de memória da sessão foi excedida.</tr></td><tr><td>A pasta de trabalho está em um estado inválido, e a operação não pode ser executada.</tr></td><tr><td>A sessão expirou devido a inatividade, e o usuário precisa recarregar a pasta de trabalho.</tr></td><tr><td>A quantidade máxima de sessões permitida por usuário foi excedida.</tr></td><tr><td>A operação foi cancelada pelo usuário.</tr></td><tr><td>Não foi possível concluir a operação porque ela está demorando muito.</tr></td><tr><td>Não foi possível concluir a solicitação, e ela deve ser repetida.</tr></td><tr><td>O período de avaliação do produto expirou.</tr></td><tr><td>A sessão expirou devido a inatividade.</tr></td><tr><td>O usuário não tem permissão para executar a operação no intervalo especificado.</tr></td><tr><td>As configurações regionais do usuário não correspondem às da sessão atual de colaboração.</tr></td><tr><td>O usuário não está mais conectado e deve atualizar ou abrir novamente a pasta de trabalho.</tr></td><tr><td>O intervalo solicitado não existe na planilha.</tr></td><tr><td>O usuário não tem permissão para editar a pasta de trabalho.</tr></td><tr><td>Não foi possível editar a pasta de trabalho porque ela está bloqueada.</tr></td><tr><td>A sessão não pode salvar a pasta de trabalho automaticamente.</tr></td><tr><td>A sessão não pode atualizar seu bloqueio no arquivo de pasta de trabalho.</tr></td><tr><td>Não foi possível processar a solicitação, e ela deve ser repetida.</tr></td><tr><td>Não foi possível verificar as informações de entrada do usuário , e elas precisam ser inseridas novamente.</tr></td><tr><td>O usuário teve o acesso negado.</tr></td><tr><td>A pasta de trabalho compartilhada precisa ser atualizada.</tr></td></table>|
|5002|Permissão negada|A operação solicitada não é permitida no modo de documento atual.|O desenvolvedor da solução envia uma operação de definição, mas o documento está em um modo que não permite alterações, como “Restringir Edição”.|
|5003|Erro de registro de eventos|Não há suporte para o tipo de evento especificado pelo objeto atual.|O desenvolvedor da solução tenta registrar ou cancelar o registro de um manipulador para um evento que não existe.|
|5004|Chamada de API inválida|Chamada à API inválida no contexto atual.|Uma chamada inválida é feita para o contexto, por exemplo, tentando usar um `CustomXMLPart` objeto no Excel.|
|5005|Dados obsoletos|Falha na operação devido aos dados estarem obsoletos no servidor.|Os dados no servidor precisam ser atualizados.|
|5006|Tempo Limite da Sessão|O tempo limite da sessão do documento esgotou-se. Recarregue o documento. |A sessão expirou.|
|5007|Chamada de API inválida|Não há suporte à enumeração no contexto atual.|Não há suporte à enumeração no contexto atual.|
|5009|Permissão negada|Acesso negado|O suplemento não tem permissão para chamar a API específica.|
|5012|Sessão inválida ou esgotada|A sessão do navegador do Office expirou ou é inválida. Para continuar, atualize a página.|A sessão entre o cliente do Office e o servidor expirou, ou então a data, hora ou fuso horário estão incorretos em seu computador.|
|6000|Nó inválido|O nó especificado não foi encontrado.|O `CustomXmlPart` nó não foi encontrado.|
|6100|Erro no XML personalizado|Erro no XML personalizado|Chamada de API inválida.|
|7000|ID inválida|A ID especificada não existe.|ID inválida.|
|7001|Navegação inválida|O objeto está localizado em um local em onde a navegação não é suportada.|The user can find the object, but cannot navigate to it. (For example, in Word, the binding is to the header, footer, or a comment.)|
|7002|Navegação inválida|O objeto está bloqueado ou protegido.|O usuário está tentando navegar até um intervalo bloqueado ou protegido.|
|7004|Navegação inválida|A operação falhou porque o índice está fora do intervalo.|O usuário está tentando navegar até um índice fora do intervalo.|
|8000|Parâmetro Ausente|We couldn't format the table cell because some parameter values are missing. Double-check the parameters and try again.|The cellFormat method is missing some parameters. For example, there are missing cells, format, or tableOptions parameters.|
|8010|Valor inválido|One or more of the cells parameters have values that aren't allowed. Double-check the values and try again.|The common cells reference enumeration is not defined. For example, All, Data, Headers.|
|8011|Valor inválido|One or more of the tableOptions parameters have values that aren't allowed. Double-check the values and try again.|Um dos valores em tableOptions é inválido.|
|8012|Valor inválido|One or more of the format parameters have values that aren't allowed. Double-check the values and try again.|Um dos valores no formato é inválido.|
|8020|Fora do intervalo|The row index value is out of the allowed range. Use a positive value (0 or higher) that's less than the number of rows.|O índice de linha é superior ao maior índice de linha da tabela ou menor do que 0.|
|8021|Fora do intervalo|The column index value is out of the allowed range. Use a positive value (0 or higher) that's less than the number of columns.|O índice de coluna é superior ao maior índice de coluna da tabela ou menor do que 0.|
|8022|Fora do intervalo|O valor está fora do intervalo permitido.|Alguns dos valores no formato estão fora dos intervalos suportados.|
|9016|Permissão negada|Permissão negada|Acesso negado.|
|9020|Erro de resposta genérica|Ocorreu um erro interno.|Refere-se a uma condição de erro interna, que pode ocorrer por vários motivos.|
|9021|Salvar Erro|Erro de conexão ao tentar salvar o item no servidor.|Não foi possível salvar o item. Isso pode ocorrer devido a um erro de conexão de servidor se estiver usando o Modo Online na área de trabalho do Outlook ou devido a uma tentativa de salvar novamente um item de rascunho que foi excluído do servidor Exchange.|
|9022|Mensagem em erro de repositório diferente|A ID do EWS não pode ser recuperada porque a mensagem é salva em outro repositório.|A ID do EWS para a mensagem atual não pôde ser recuperada, pois a mensagem pode ter sido movida ou a caixa de correio de envio pode ter sido alterada.|
|9041|Erro de rede|O usuário não está mais conectado à rede. Verifique sua conexão de rede e tente novamente.|O usuário não tem mais acesso à rede ou à Internet.|
|9043|Tipo de anexo sem suporte|Não há suporte para o tipo de anexo.|A API não dá suporte ao tipo de anexo. Por exemplo, `item.getAttachmentContentAsync` gera esse erro se o anexo for uma imagem inserida no Formato Rich Text ou se for um tipo de item diferente de um email ou item de calendário (como um contato ou item de tarefa).|
|12002|*Não aplicável.*|*Não aplicável.*|Uma destas opções:<br> - Não existe uma página na URL transmitida para `displayDialogAsync`.<br> - A página transmitida para `displayDialogAsync` foi carregada, mas a caixa de diálogo foi direcionada para uma página que ela não consegue localizar nem carregar ou foi direcionada para uma URL com sintaxe inválida. Lançado dentro da caixa de diálogo e dispara um evento `DialogEventReceived` na página de host.|
|12003|*Não aplicável.*|*Não aplicável.*|A caixa de diálogo foi direcionada para uma URL com o protocolo HTTP. HTTPS é necessário. Lançado dentro da caixa de diálogo e dispara um evento `DialogEventReceived` na página de host.|
|12004|*Não aplicável.*|*Não aplicável.*|O domínio que a URL transmitiu para `displayDialogAsync` não é confiável. O domínio deve ser o mesmo domínio que o da página de host (incluindo o protocolo e o número da porta). Lançada por chamada de `displayDialogAsync`.|
|12005|*Não aplicável.*|*Não aplicável.*|A URL passada para `displayDialogAsync` usa o protocolo HTTP. HTTPS é necessário. Lançada por chamada de `displayDialogAsync`. (Em algumas versões do Office, a mensagem de erro retornada com 12005 é a mesma retornada para 12004.)|
|12006|*Não aplicável.*|*Não aplicável.*|A caixa de diálogo foi fechada, geralmente pelo usuário ter escolhido o botão **X**. Lançado dentro da caixa de diálogo e dispara um evento `DialogEventReceived` na página de host.|
|12007|*Não aplicável.*|*Não aplicável.*|Uma caixa de diálogo já está aberta na janela do host. Uma janela do host, como um painel de tarefas, só pode ter uma caixa de diálogo aberta por vez. Lançada por chamada de `displayDialogAsync`.|
|12009|*Não aplicável.*|*Não aplicável.*|O usuário opta por ignorar a caixa de diálogo. Este erro pode ocorrer em versões online do Office, em que os usuários podem optar por não permitir que um suplemento apresente uma caixa de diálogo. Lançada por chamada de `displayDialogAsync`.|
|12011|*Não aplicável.*|*Não aplicável.*|O navegador do usuário é configurado de uma maneira que bloqueia pop-ups. Esse erro pode ocorrer no Office na Web se o navegador for Safari e ele estiver configurado para bloquear pop-ups ou se o navegador for o Edge Legacy e o domínio do suplemento estiver em uma zona de segurança diferente do domínio que a caixa de diálogo está tentando abrir. Lançada por chamada de `displayDialogAsync`.|
|13nnn|*Não aplicável.*|*Não aplicável.*|Consulte [Causas e tratamento de erros de getAccessToken](../develop/troubleshoot-sso-in-office-add-ins.md#causes-and-handling-of-errors-from-getaccesstoken).|

## <a name="binding-creation-error-conditions"></a>Condições do erro de criação de associação

When a binding is created in the API, indicate the binding type that you want to use. The following tables lists the binding types and the resulting binding behaviors that are expected.

### <a name="behavior-in-excel"></a>Comportamento no Excel

A tabela a seguir resume o comportamento de associação no Excel.

|Tipo de associação especificado|Seleção real|Comportamento|
|:-----|:-----|:-----|
|Matriz|Intervalo de células (incluindo dentro de uma tabela e uma única célula)|Uma associação de tipo `matrix` é criada nas células selecionadas. Nenhuma modificação no documento é esperada.|
|Matriz|Texto selecionado na célula|Uma associação de tipo `matrix` é criada em toda a célula. Nenhuma modificação no documento é esperada.|
|Matriz|Várias seleções/seleção inválida (por exemplo, o usuário seleciona uma imagem, objeto ou WordArt.)|Não é possível criar a associação.|
|Tabela|Intervalo de células (inclui uma única célula)|Não é possível criar a associação.|
|Tabela|Intervalo de células dentro de uma tabela (inclui uma única célula dentro de uma tabela, ou a tabela inteira ou texto dentro de uma célula em uma tabela)|Uma associação é criada na tabela inteira.|
|Tabela|Metade da seleção em uma tabela e metade da seleção fora da tabela|Não é possível criar a associação.|
|Tabela|Texto selecionado na célula (não na tabela).|Não é possível criar a associação.|
|Tabela|Várias seleções/seleção inválida (por exemplo, o usuário seleciona uma imagem, objeto, WordArt etc.)|Não é possível criar a associação.|
|Texto|Intervalo de células|Não é possível criar a associação.|
|Texto|Intervalo de células dentro de uma tabela|Não é possível criar a associação.|
|Texto|Célula única|Uma associação de tipo `text` é criada.|
|Texto|Célula única dentro de uma tabela|Uma associação de tipo `text` é criada.|
|Texto|Texto selecionado na célula|Uma associação de tipo `text` em toda a célula é criada.|

### <a name="behavior-in-word"></a>Comportamento no Word

A tabela a seguir resume o comportamento de associação no Word.

|Tipo de associação especificado|Seleção real|Comportamento|
|:-----|:-----|:-----|
|Matriz|Texto|Não é possível criar a associação.|
|Matriz|Tabela inteira|Uma associação de tipo `matrix` é criada. O documento é alterado e um controle de conteúdo deve encapsular a tabela. |
|Matriz|Intervalo dentro de uma tabela|Não é possível criar a associação.|
|Matriz|Seleção inválida (por exemplo, múltiplos objetos, objetos inválidos etc.)|Não é possível criar a associação.|
|Tabela|Texto|Não é possível criar a associação.|
|Tabela|Tabela inteira|Uma associação de tipo `text` é criada.|
|Table|Intervalo dentro de uma tabela|Não é possível criar a associação.|
|Tabela|Seleção inválida (por exemplo, múltiplos objetos, objetos inválidos etc.)|Não é possível criar a associação.|
|Texto|Tabela inteira|Uma associação de tipo `text` é criada.|
|Texto|Intervalo dentro de uma tabela|Não é possível criar a associação.|
|Texto|Seleção múltipla|A última seleção será envolvida com um controle de conteúdo e uma associação a esse controle. Um controle de conteúdo do tipo `text` é criado.|
|Texto|Seleção inválida (por exemplo, múltiplos objetos, objetos inválidos etc.)|Não é possível criar a associação.|

## <a name="see-also"></a>Confira também

- [Ciclo de vida de desenvolvimento de suplementos do Office](../overview/office-add-ins.md)
- [Entendendo a API de JavaScript do Office](../develop/understanding-the-javascript-api-for-office.md)
- [Tratamento de erro com as APIs JavaScript específicas do aplicativo](../testing/application-specific-api-error-handling.md)
- [Solucionar problemas de mensagens de erro no logon único (SSO)](../develop/troubleshoot-sso-in-office-add-ins.md)
- [Solucionar erros de desenvolvimento com Suplementos do Office](../testing/troubleshoot-development-errors.md)
