# <a name="whats-changed-in-the-javascript-api-for-office"></a>O que mudou na API JavaScript para Office

A API JavaScript para Office é atualizada periodicamente com objetos, métodos, propriedades, eventos e enumerações novos e atualizados para estender a funcionalidade dos seus suplementos do Office. Use os links abaixo para ver os membros da API novos e atualizados.

Para desenvolver suplementos usando novos membros da API, você precisa [atualizar os arquivos da API JavaScript para Office em seu projeto](/office/dev/add-ins/develop/update-your-javascript-api-for-office-and-manifest-schema-version).

Para ver todos os membros da API, incluindo aqueles que não sofreram alterações nas atualizações anteriores, confira [API JavaScript para Office](javascript-api-for-office.md).

## <a name="new-and-updated-apis"></a>APIs novas e atualizadas

### <a name="new-and-updated-objects"></a>Objetos novos e atualizados

|**Objeto**|**Descrição**|**Versão adicionada ou atualizada**|
|:-----|:-----|:-----|
|`Item`|Atualizações e adições para:<br><ul><li><p>Os métodos `getSelectedDataAsync` e `setSelectedDataAsync` para dar suporte à seleção do usuário e substituí-lo no assunto e no corpo de uma mensagem ou um compromisso.</p></li><li><p>Os métodos `displayReplyAllForm` e `displayReplyForm` para dar suporte à adição de um anexo ao formulário de resposta de um compromisso.</p></li></ul>|Mailbox 1.2|
|`Item`|Atualizado para incluir campos e métodos para criação de suplementos do Outlook no modo de composição. |1.1|
|`Binding`|Atualizado para dar suporte à associação de tabela em suplementos de conteúdo para o Access.|1.1|
|`Bindings`|Atualizado para dar suporte à associação de tabela em suplementos de conteúdo para o Access.|1.1|
|`Body`|Adicionado para permitir a criação e a edição do corpo de uma mensagem ou compromisso em suplementos do Outlook no modo de composição.|1.1|
|`Document`|Atualizações e adições para: <ul><li><p>Suporte às propriedades <a href="/javascript/api/office/office.document" target="_blank">mode</a>, <a href="/javascript/api/office/office.document#settings" target="_blank">settings</a> e <a href="/javascript/api/office/office.document" target="_blank">url</a> em suplementos de conteúdo para o Access.</p></li><li><p>Navegar até locais e objetos dentro do documento com o método <a href="/javascript/api/office/office.document#getfileasync-filetype--options--callback-" target="_blank">goToByIdAsync</a> em suplementos para Excel e PowerPoint.</p></li><li><p>Obter as propriedades de arquivo com o método <a href="/javascript/api/office/office.document#getfilepropertiesasync-options--callback-" target="_blank">getFileProperties</a> em suplementos para Excel, PowerPoint e Word.</p></li><li><p>Navegar até locais e objetos dentro do documento com o método <a href="/javascript/api/office/office.document#gotobyidasync-id--gototype--options--callback-" target="_blank">goToByIdAsync</a> em suplementos para Excel e PowerPoint.</p></li><li><p>Obter a identificação, o título e o índice dos slides selecionados com o método <a href="/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-" target="_blank">getSelectedDataAsync</a> (ao especificar a nova enumeração <span class="keyword">Office.CoercionType.SlideRange</span><a href="/javascript/api/office/office.coerciontype" target="_blank">coercionType</a>) em suplementos do PowerPoint.</p></li></ul>|1.1|
|`Location`|Adicionado a fim de permitir a configuração do local de um compromisso em suplementos do Outlook no modo de composição.|1.1|
|`Office`|Método de seleção atualizado para oferecer suporte à obtenção de associações em suplementos de conteúdo para Access.|1.1|
|`Recipients`|Adicionado para permitir a obtenção e a configuração dos destinatários de uma mensagem ou de um compromisso no modo de composição.|1.1|
|`Settings`|Atualizado para oferecer suporte à criação de configurações personalizadas em suplementos de conteúdo para o Access.|1.1|
|`Subject`|Adicionado para permitir a obtenção e a configuração do assunto de uma mensagem ou de um compromisso em suplementos do Outlook no modo de composição.|1.1|
|`Time`|Adicionado para permitir a obtenção e a configuração das horas de início e de término de um compromisso em suplementos do Outlook no modo de composição.|1.1|

### <a name="new-and-updated-enumerations"></a>Descrição

|**Objeto**|**ActiveView**|**Versão**|
|:-----|:-----|:-----|
|`ActiveView`|Adicionado para que os suplementos do PowerPoint possam determinar se os usuários estão exibindo a apresentação (**Apresentação de Slides**) ou editando slides. |1.1|
|`CoercionType`|Atualizado com **Office.CoercionType.SlideRange** para oferecer suporte à obtenção do intervalo de slides selecionado com o método **getSelectedDataAsync** em suplementos para PowerPoint.|1.1|
|`EventType`|Atualizado para incluir o novo evento ActiveViewChanged.|1.1|
|`FileType`|Atualizado para especificar a saída no formato PDF.|1.1|
|`GoToType`|Adicionado para especificar o local ou o objeto a ser acessado no documento.|1.1|

