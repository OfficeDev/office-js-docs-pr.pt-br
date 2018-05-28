---
title: Solicitar permiss?es para uso da API em suplementos do painel de tarefas e conte?do
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: c73ddbaa3d517f82b5fdf815b7e86f4e7a91a541
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="requesting-permissions-for-api-use-in-content-and-task-pane-add-ins"></a>Solicitar permiss?es para uso da API em suplementos do painel de tarefas e conte?do

Este artigo descreve os diferentes n?veis de permiss?o que voc? pode declarar no manifesto do suplemento de conte?do ou de painel de tarefas para especificar o n?vel de acesso da API JavaScript que o suplemento requer para seus recursos. 




## <a name="permissions-model"></a>Modelo de permiss?es


Um modelo de permiss?es de acesso da API JavaScript com cinco n?veis fornece a base para a privacidade e a seguran?a dos usu?rios dos suplementos de conte?do e de painel de tarefas. A Figura 1 mostra os cinco n?veis de permiss?es da API que voc? pode declarar no manifesto do suplemento.


*Figura 1. Modelo de permiss?o com cinco n?veis para os suplementos do conte?do e do painel de tarefas*

![N?veis de permiss?es para aplicativos do painel de tarefas](../images/office15-app-sdk-task-pane-app-permission.png)



Essas permiss?es especificam o subconjunto da API que o tempo de execu??o do suplemento permitir? que o suplemento de conte?do ou de painel de tarefas use quando um usu?rio inserir e ativar o suplemento (confiar nele). Para declarar o n?vel de permiss?o que o suplemento do conte?do ou do painel de tarefas requer, especifique um dos valores de texto da permiss?o no elemento [Permissions](http://msdn.microsoft.com/en-us/library/d4cfe645-353d-8240-8495-f76fb36602fe%28Office.15%29.aspx) do manifesto do suplemento. O exemplo a seguir solicita a permiss?o  **WriteDocument**, que autorizar? somente os m?todos que podem gravar no documento, mas n?o l?-lo.




```XML
<Permissions>WriteDocument</Permissions>
```

Como pr?tica recomendada, voc? deve solicitar permiss?es com base no princ?pio do _menor privil?gio_. Ou seja, voc? deve solicitar permiss?o para acessar apenas o subconjunto m?nimo da API que o suplemento requer para funcionar corretamente. Por exemplo, se o suplemento precisar apenas ler os dados no documento de um usu?rio para seus recursos, voc? n?o deve solicitar mais do que a permiss?o **ReadDocument**.

A tabela a seguir descreve o subconjunto da API JavaScript que ? habilitado por cada n?vel de permiss?o.



|**Permiss?o**|**Subconjunto habilitado da API**|
|:-----|:-----|
|**Restrito**|Os m?todos do objeto [Settings](https://dev.office.com/reference/add-ins/shared/settings) e o m?todo [Document.getActiveViewAsync](https://dev.office.com/reference/add-ins/shared/document.getactiveviewasync). Esse ? o n?vel de permiss?o m?nimo que pode ser solicitado por um suplemento de conte?do ou de painel de tarefas.|
|**ReadDocument**|Al?m da API autorizada pela permiss?o **Restricted**, adiciona acesso aos membros da API necess?rios para ler o documento e gerenciar associa??es. Isso inclui o uso de:<br/><ul><li>O m?todo <a href="https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync" target="_blank">Document.getSelectedDataAsync</a> para obter o texto selecionado, HTML (Word apenas) ou os dados tabulares, mas n?o o c?digo do Open Office XML (OOXML) subjacente que cont?m todos os dados no documento.</p></li><li><p>M?todo <a href="https://dev.office.com/reference/add-ins/shared/document.getfileasync" target="_blank">Document.getFileAsync</a> para acessar todo o texto no documento, mas n?o a c?pia bin?ria OOXML subjacente do documento.</p></li><li><p>M?todo <a href="http://msdn.microsoft.com/en-us/library/5372ffd8-579d-4fcb-9e5b-e9a2128f3201(Office.15).aspx" target="_blank">Binding.getDataAsync</a> para a leitura dos dados associados no documento.</p></li><li><p>M?todos <a href="http://msdn.microsoft.com/en-us/library/afbadac7-60c7-47cb-9477-6e9466ded44c(Office.15).aspx" target="_blank">addFromNamedItemAsync</a>, <a href="http://msdn.microsoft.com/en-us/library/9dc03608-b08b-4700-8be1-3c86ae236799(Office.15).aspx" target="_blank">addFromPromptAsync</a>, <a href="http://msdn.microsoft.com/en-us/library/edc99214-e63e-43f2-9392-97ead42fc155(Office.15).aspx" target="_blank">addFromSelectionAsync</a> do objeto <span class="keyword">Bindings</span> para criar associa??es no documento.</p></li><li><p>M?todos <a href="http://msdn.microsoft.com/en-us/library/ef902b73-cc4c-4551-95de-d8a51eeba82f(Office.15).aspx" target="_blank">getAllAsync</a>, <a href="http://msdn.microsoft.com/en-us/library/2727c891-bc05-465c-9324-113fbfeb3fbb(Office.15).aspx" target="_blank">getByIdAsync</a> e <a href="http://msdn.microsoft.com/en-us/library/ad285984-8b44-435d-9b84-f0ade570c896(Office.15).aspx" target="_blank">releaseByIdAsync</a> do objeto <span class="keyword">Bindings</span> para acessar e remover as associa??es no documento.</p></li><li><p>M?todo <a href="http://msdn.microsoft.com/en-us/library/2533a563-95ae-4d52-b2d5-a6783e4ef5b4(Office.15).aspx" target="_blank">Document.getFilePropertiesAsync</a> para acessar as propriedades de arquivo do documento, como a URL do documento.</p></li><li><p>M?todo <a href="http://msdn.microsoft.com/en-us/library/35dda81c-235e-4eab-8a77-9acb3b73a380(Office.15).aspx" target="_blank">Document.goToByIdAsync</a> para navegar at? os objetos nomeados e locais no documento.</p></li><li><p>Para os suplementos do painel de tarefas do Project, todos os m?todos "get" do objeto <a href="http://msdn.microsoft.com/en-us/library/1908af4f-93b9-4859-87e3-06942014fae1(Office.15).aspx" target="_blank">ProjectDocument</a>. </p></li></ul>|
|**ReadAllDocument**|Al?m da API autorizada pelas permiss?es **Restricted** e **ReadDocument**, permite os seguintes acessos adicionais aos dados do documento:<br/><ul><li><p>Os m?todos <span class="keyword">Document.getSelectedDataAsync</span> e <span class="keyword">Document.getFileAsync</span> podem acessar o c?digo OOXML subjacente do documento (que, al?m de texto, pode conter formata??o, links, gr?ficos incorporados, coment?rios, revis?es, etc.).</p></li></ul>|
|**WriteDocument**|Al?m da API autorizada pela permiss?o **Restricted**, adiciona acesso aos seguintes membros da API:<br/><ul><li><p>M?todo <a href="http://msdn.microsoft.com/en-us/library/998f38dc-83bd-4659-a759-4758c632a6ef(Office.15).aspx" target="_blank">Document.setSelectedDataAsync</a> para gravar na sele??o do usu?rio no documento.</p></li></ul>|
|**ReadWriteDocument**|Al?m da API autorizada pelas permiss?es **Restricted**, **ReadDocument**, **ReadAllDocument** e **WriteDocument**, cont?m acesso a todas as API remanescentes compat?veis com suplementos de painel de tarefas e de conte?do, incluindo m?todos para se inscrever em eventos. Voc? precisa declarar a permiss?o **ReadWriteDocument** para acessar esses membros adicionais da API:<br/><ul><li><p>M?todo <a href="http://msdn.microsoft.com/en-us/library/6a59bb6d-40b6-4a95-9b98-d70d4616de09(Office.15).aspx" target="_blank">Binding.setDataAsync</a> para gravar nas regi?es associadas do documento.</p></li><li><p>M?todo <a href="http://msdn.microsoft.com/en-us/library/1cd23454-8435-4e13-98b3-d0d29ed278a8(Office.15).aspx" target="_blank">TableBinding.addRowsAsync</a> para adicionar linhas ?s tabelas associadas.</p></li><li><p>M?todo <a href="http://msdn.microsoft.com/en-us/library/8f1bfa81-3850-4ea1-ba2e-c9bcf5847a44(Office.15).aspx" target="_blank">TableBinding.addColumnsAsync</a> para adicionar colunas ?s tabelas associadas.</p></li><li><p>M?todo <a href="http://msdn.microsoft.com/en-us/library/8f5cc783-384d-4520-a218-190dfed74dd2(Office.15).aspx" target="_blank">TableBinding.deleteAllDataValuesAsync</a> para excluir todos os dados em uma tabela associada.</p></li><li><p>M?todos <a href="http://msdn.microsoft.com/en-us/library/49712906-f582-4055-9ef8-6edde6e97679(Office.15).aspx" target="_blank">setFormatsAsync</a>, <a href="http://msdn.microsoft.com/en-us/library/cc56e9c0-b33c-4d9b-b676-a7e50f757c10(Office.15).aspx" target="_blank">clearFormatsAsync</a> e <a href="http://msdn.microsoft.com/en-us/library/2885fc57-4527-4ca4-a43d-9ee447ec27d3(Office.15).aspx" target="_blank">setTableOptionsAsync</a> do objeto <span class="keyword">TableBinding</span> para definir a formata??o e as op??es nas tabelas associadas.</p></li><li><p>Todos os membros dos objetos <a href="http://msdn.microsoft.com/en-us/library/dc1518de-47fa-4108-aab7-04a022724b04(Office.15).aspx" target="_blank">CustomXmlNode</a>, <a href="http://msdn.microsoft.com/en-us/library/83f0e668-8236-4f2f-a20f-b173a9e3f65f(Office.15).aspx" target="_blank">CustomXmlPart</a>, <a href="http://msdn.microsoft.com/en-us/library/ba40cd4c-29bb-4f31-875d-6f1382fd1ee8(Office.15).aspx" target="_blank">CustomXmlParts</a> e <a href="http://msdn.microsoft.com/en-us/library/18b9aa8c-83e7-4c2f-8530-6a0ac8ce5535(Office.15).aspx" target="_blank">CustomXmlPrefixMappings</a>.</p></li><li><p>Todos os m?todos para se inscrever em eventos compat?veis com suplementos de conte?do e de painel de tarefas, especificamente os m?todos <span class="keyword">addHandlerAsync</span> e <span class="keyword">removeHandlerAsync</span> dos objetos <a href="http://msdn.microsoft.com/en-us/library/42882642-d22b-47d2-a8d3-3aa8c6a4435e(Office.15).aspx" target="_blank">Binding</a>, <a href="http://msdn.microsoft.com/en-us/library/83f0e668-8236-4f2f-a20f-b173a9e3f65f(Office.15).aspx" target="_blank">CustomXmlPart</a>, <a href="http://msdn.microsoft.com/en-us/library/f8859516-cc1f-4b20-a8f3-cee37a983e70(Office.15).aspx" target="_blank">Document</a>, <a href="http://msdn.microsoft.com/en-us/library/1908af4f-93b9-4859-87e3-06942014fae1(Office.15).aspx" target="_blank">ProjectDocument</a> e <a href="http://msdn.microsoft.com/en-us/library/ad733387-a58c-4514-8fc2-53e64fad468d(Office.15).aspx" target="_blank">Settings</a>.</p></li></ul>|

## <a name="see-also"></a>Veja tamb?m

- [Privacidade e seguran?a para Suplementos do Office](../concepts/privacy-and-security.md)
    


