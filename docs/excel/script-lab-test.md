# <a name="testing-script-lab-integration"></a>Teste de integração de laboratório de Script

Este é um arquivo de teste de exemplo, destinado a demonstrar um recurso ScriptLab futuro que permitirá que os desenvolvedores experimentem trechos no Excel, no Word e no PowerPoint.  

## <a name="pre-reqs"></a>Pré-requisitos:
- Você precisará de uma URL de visualização de um trecho de ScriptLab
- Observação: *Devemos* indicar que ScriptLab precisa do Office 365 para explorar os trechos mais recentes.  Os desenvolvedores podem obter uma assinatura do Office 365 por meio de nosso [Programa de desenvolvedor do Office 365](https://dev.office.com/devprogram), apenas para fins de desenvolvimento.  


## <a name="try-it-out-button"></a>'Botão' Experimente
Dessa forma, adicionaremos um botão Experimentar, que é recomendável associar a um trecho de código.  Para habilitar isso, estamos usando uma classe do Office UI Fabric para definir o estilo de um link como um botão. No link em si, lembre-se de configurar o atributo *aria label*.

**Demonstração:**

<a href="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">Experimente</a>


<button href="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">Experimente</button>


**Código:**
```html
<a href="ahttps://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">Try it out</a>
```



## <a name="embed-script-lab-as-an-iframe"></a>Incorporar laboratório de script como um iframe
Nesse modo, incorporaremos um trecho diretamente como um iframe nos documentos. A largura foi definida como 95% (com base na largura de todos os outros trechos), e recomendamos remover o fameborder do iframe.  A altura normalmente deve ser ajustada para corresponder ao trecho de código.

**Demonstração:**
<iframe src="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" height="600px" width="95%" frameborder="0"></iframe>

**Código:**
```html
<iframe src="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" height="600px" width="95%" frameborder="0"></iframe>
```

## <a name="testing-considerations"></a>Considerações sobre os testes
Precisamos verificar assinaturas móveis não Office 365 (temos comentários sobre os documentos do office js em que muitas desenvolvedores usaram em 2013 ou versões anteriores.  

Para o caminho de inserção, precisamos de aprovação final e precisamos garantir que o conteúdo exposto na página de exibição atenda a nossas diretrizes de acessibilidade.
