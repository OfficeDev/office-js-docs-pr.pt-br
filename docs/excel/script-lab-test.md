---
title: Teste de integração de laboratório de Script
description: 'Este arquivo de teste de exemplo demonstra um recurso de ScriptLab futuro que permitirá aos desenvolvedores experimentar trechos no Excel, no Word e no PowerPoint.'
ms.date: 03/14/2018
---


# <a name="testing-script-lab-integration"></a>Teste de integração de laboratório de Script

Este arquivo de teste de exemplo demonstra um recurso de ScriptLab futuro que permitirá aos desenvolvedores experimentar trechos no Excel, no Word e no PowerPoint. 

## <a name="prerequisites"></a>Pré-requisitos

- Você precisará de uma URL de Exibição de um trecho de código do ScriptLab.

> [!NOTE] 
> *Devemos* indicar que o ScriptLab precisa do Office 365 para explorar os trechos mais recentes. Os desenvolvedores podem receber uma assinatura de desenvolvedor do Office 365 por meio de nosso [Programa de Desenvolvedor do Office 365](https://developer.microsoft.com/en-us/office/dev-program), apenas para fins de desenvolvimento. Confira as instruções passo a passo sobre como participar do Programa para Desenvolvedores do Office 365, entrar e configurar sua assinatura na [documentação do Programa para Desenvolvedores do Office 365](https://docs.microsoft.com/pt-br/office/developer-program/office-365-developer-program). 


## <a name="try-it-out-button"></a>Botão Experimentar

Dessa forma, adicionaremos um botão **Experimentar**, o qual recomendamos associar a um trecho de código. Para habilitar isso, usamos uma classe do Office UI Fabric para definir o estilo de um link como um botão. No próprio link, configure o atributo `aria label`.

### <a name="demo"></a>Demonstração

<a href="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">Experimente</a>


<button href="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">Experimente</button>


### <a name="code"></a>Código

```html
<a href="ahttps://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">Try it out</a>
```



## <a name="embed-script-lab-as-an-iframe"></a>Incorporar laboratório de script como um iframe

Nesse modo, incorporaremos um trecho diretamente como um iframe nos documentos. A largura foi definida como 95% (com base na largura de todos os outros trechos), e recomendamos remover o fameborder do iframe. A altura normalmente deve ser ajustada para corresponder ao trecho de código.

### <a name="demo"></a>Demonstração

<iframe src="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" height="600px" width="95%" frameborder="0"></iframe>

### <a name="code"></a>Código

```html
<iframe src="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" height="600px" width="95%" frameborder="0"></iframe>
```

## <a name="testing-considerations"></a>Considerações sobre os testes

Precisamos verificar assinaturas móveis que não sejam do Office 365 (recebemos comentários no office-js-docs dizendo que muitos desenvolvedores usavam a versão 2013 ou anteriores).  

Para o caminho de inserção, precisamos de aprovação final e precisamos garantir que o conteúdo exposto na página de exibição atenda a nossas diretrizes de acessibilidade.


