---
title: Criar suplementos melhores para o Word com o Office Open XML
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: ed44b9d331670ac7bf9fb625555dcd05f7bff7ec
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="create-better-add-ins-for-word-with-office-open-xml"></a>Criar suplementos melhores para o Word com o Office Open XML

**Fornecido por:** Stephanie Krieger, Microsoft Corporation | Juan Balmori Labra, Microsoft Corporation

Se voc? est? criando suplementos do Office para serem executados no Word, talvez j? saiba que a API JavaScript para Office (Office.js) oferece v?rios formatos para ler e gravar o conte?do de documentos. Eles s?o chamados de tipos de coer??o e incluem texto sem formata??o, tabelas, HTML e Office Open XML.

Ent?o, quais s?o suas op??es quando voc? precisa adicionar conte?do avan?ado a um documento, como imagens, tabelas formatadas, gr?ficos ou apenas texto formatado? Voc? pode usar HTML para inserir alguns tipos de conte?do avan?ado, como imagens. Dependendo do cen?rio, pode haver desvantagens na coer??o de HTML, como limita??es nas op??es de formata??o e posicionamento dispon?veis para o conte?do. Como o Office Open XML ? a linguagem na qual os documentos do Word (como .docx e .dotx) s?o gravados, voc? pode inserir praticamente qualquer tipo de conte?do que um usu?rio pode adicionar a um documento do Word, com praticamente qualquer tipo de formata??o que o usu?rio possa aplicar. Determinar a marca??o do Office Open XML necess?ria para fazer isso ? mais f?cil do que voc? imagina.

> [!NOTE]
> O Office Open XML tamb?m ? a linguagem por tr?s dos documentos do PowerPoint e do Excel (e, a partir do Office 2013, do Visio). No entanto, atualmente, voc? pode fazer a coer??o de conte?do como Office Open XML somente em Suplementos do Office criados para o Word. Para saber mais sobre o Office Open XML, incluindo a documenta??o de refer?ncia completa da linguagem, confira [Recursos adicionais](#see-also).

Para come?ar, veja alguns dos tipos de conte?do que voc? pode inserir usando a coer??o do Office Open XML. Baixe o exemplo de c?digo [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML), que cont?m a marca??o do Office Open XML e o c?digo Office.js necess?rio para inserir qualquer um dos exemplos a seguir no Word.

> [!NOTE]
> Ao longo deste artigo, os termos **tipos de conte?do** e **conte?do avan?ado** referem-se aos tipos de conte?do avan?ado que voc? pode inserir em um documento do Word.


*Figura 1. Texto com formata??o direta*


![Texto com formata??o direta aplicada.](../images/office15-app-create-wd-app-using-ooxml-fig01.png)

Voc? pode usar a formata??o direta para especificar a apar?ncia exata que o texto ter?, independentemente da formata??o existente no documento do usu?rio.

*Figura 2. Texto formatado com um estilo*


![Texto formatado com estilo de par?grafo.](../images/office15-app-create-wd-app-using-ooxml-fig02.png)

Voc? pode usar um estilo para coordenar automaticamente a apar?ncia do texto que insere com o documento do usu?rio.

*Figura 3. Uma imagem simples*


![Imagem de um logotipo.](../images/office15-app-create-wd-app-using-ooxml-fig03.png)

Voc? pode usar o mesmo m?todo para inserir qualquer formato de imagem compat?vel com o Office.

*Figura 4. Uma imagem formatada usando efeitos e estilos de imagem*


![Imagem formatada no Word 2013.](../images/office15-app-create-wd-app-using-ooxml-fig04.png)


A adi??o de efeitos e formata??o de alta qualidade ?s imagens requer muito menos marca??o do que voc? poderia esperar.

*Figura 5. Um controle de conte?do*


![Texto em um controle de conte?do vinculado.](../images/office15-app-create-wd-app-using-ooxml-fig05.png)

Voc? pode usar controles de conte?do com o suplemento para adicionar conte?do em um local especificado (associado) em vez de na sele??o.

*Figura 6. Uma caixa de texto com formata??o do WordArt*


![Texto formatado com efeitos de texto WordArt.](../images/office15-app-create-wd-app-using-ooxml-fig06.png)

Os efeitos de texto est?o dispon?veis no Word para o texto dentro de uma caixa de texto (como mostrado aqui) ou para o corpo do texto normal.

*Figura 7. Uma forma*


![Uma forma de desenho do Office 2013 no Word 2013.](../images/office15-app-create-wd-app-using-ooxml-fig07.png)

Voc? pode inserir formas de desenho internas ou personalizadas, com ou sem texto e efeitos de formata??o.

*Figura 8. Uma tabela com formata??o direta*


![Uma tabela formatada no Word 2013.](../images/office15-app-create-wd-app-using-ooxml-fig08.png)

Voc? pode incluir formata??o de texto, bordas, sombreamento, dimensionamento de c?lulas ou qualquer formata??o de tabela que seja necess?ria.

*Figura 9. Uma tabela formatada usando um estilo de tabela*


![Uma tabela formatada no Word 2013.](../images/office15-app-create-wd-app-using-ooxml-fig09.png)

Voc? pode usar estilos de tabela internos ou personalizados com a mesma facilidade com que usa um estilo de par?grafo para o texto.

*Figura 10. Um diagrama do SmartArt*


![Um diagrama SmartArt din?mico no Word 2013.](../images/office15-app-create-wd-app-using-ooxml-fig10.png)

O Office 2013 oferece uma ampla variedade de layouts de diagrama do SmartArt (e voc? pode usar o Office Open XML para criar os seus pr?prios).

*Figura 11. Um gr?fico*


![Um gr?fico no Word 2013.](../images/office15-app-create-wd-app-using-ooxml-fig11.png)

Voc? pode inserir gr?ficos do Excel como gr?ficos din?micos em documentos do Word, o que tamb?m significa que voc? pode us?-los no seu suplemento do Word. Como voc? pode ver pelos exemplos anteriores, ? poss?vel usar a coer??o do Office Open XML para inserir praticamente qualquer tipo de conte?do que um usu?rio pode inserir em seu pr?prio documento. H? duas maneiras simples de obter a marca??o do Office Open XML necess?ria. Adicionar conte?do avan?ado a um documento do Word 2013 em branco e salvar o arquivo no formato de Documento XML do Word ou usar um suplemento de teste com o m?todo [getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) para obter a marca??o. As duas abordagens fornecem basicamente o mesmo resultado.

    
> [!NOTE]
> Um documento do Office Open XML ? realmente um pacote compactado de arquivos que representa o conte?do do documento. Salvar o arquivo no formato de Documento XML do Word lhe fornece todo o pacote do Office Open XML compactado em um arquivo XML, que tamb?m ? o que voc? obt?m ao usar **getSelectedDataAsync** para recuperar a marca??o XML do Office Open XML.

Se voc? salvar o arquivo em um formato XML do Word, observe que h? duas op??es na lista Salvar como Tipo na caixa de di?logo Salvar como para arquivos no formato .xml. Certifique-se de escolher **Documento XML do Word** e n?o a op??o do Word 2003. Baixe o c?digo de exemplo nomeado [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML), que pode ser usado como uma ferramenta para recuperar e testar sua marca??o. Ent?o ? s? isso que preciso fazer? Bem, n?o exatamente. Sim, para muitos cen?rios, voc? poderia usar todo o resultado compactado do Office Open XML que obt?m com um dos m?todos anteriores, e ele funcionaria. A boa not?cia ? que voc? provavelmente n?o precisa da maioria dessa marca??o. Se voc? ? um dos muitos desenvolvedores de suplementos que est?o vendo a marca??o do Office Open XML pela primeira vez, tentar entender a grande quantidade de marca??o obtida at? para o conte?do mais simples pode parecer assustador, mas n?o precisa ser assim. Neste t?pico, usaremos alguns cen?rios comuns que obtivemos da comunidade de desenvolvedores de Suplementos do Office para mostrar t?cnicas que simplificam o Office Open XML para uso em suplementos. Exploraremos a marca??o para alguns tipos de conte?do mostrados anteriormente, al?m das informa??es necess?rias para minimizar a carga do Office Open XML. Tamb?m examinaremos o c?digo necess?rio para inserir conte?do avan?ado em um documento na sele??o ativa e a maneira de usar o Office Open XML com o objeto de associa??o para adicionar ou substituir conte?do em locais espec?ficos.

## <a name="exploring-the-office-open-xml-document-package"></a>Explorar o pacote de documento do Office Open XML


Ao usar [getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) para recuperar o Office Open XML para uma sele??o de conte?do (ou ao salvar o documento no formato de Documento XML do Word), o que voc? obt?m n?o ? apenas a marca??o que descreve o conte?do selecionado, ? um documento inteiro com v?rias op??es e configura??es das quais voc? certamente n?o necessita. De fato, se voc? usar esse m?todo com um documento que contenha um suplemento de painel de tarefas, a marca??o obtida incluir? at? mesmo o painel de tarefas.

At? mesmo um pacote de documento simples do Word inclui partes para propriedades de documentos, estilos, tema (configura??es de formata??o), configura??es da Web, fontes e muito mais, al?m de partes para o conte?do real.

Por exemplo, digamos que voc? queira inserir apenas um par?grafo de texto com formata??o direta, conforme mostrado anteriormente na Figura 1. Ao usar o Office Open XML para o texto formatado com **getSelectedDataAsync**, voc? v? uma grande quantidade de marca??o. A marca??o inclui um elemento de pacote que representa um documento inteiro, que cont?m v?rias partes (comumente conhecidas como partes do documento ou, no Office Open XML, partes do pacote), como pode ver listado na Figura 13. Cada parte representa um arquivo separado dentro do pacote.


> [!TIP]
> Voc? pode editar a marca??o do Office Open XML em um editor de texto como o Bloco de Notas. Se abri-lo no Visual Studio 2015, pode usar **Editar > Avan?ado > Formatar Documento** (Ctrl+K, Ctrl+D) para formatar o pacote, facilitando a edi??o. Em seguida, voc? pode recolher ou expandir partes de um documento ou se??es delas, conforme mostrado na Figura 12, para examinar e editar mais facilmente o conte?do do pacote do Office Open XML. Cada parte do documento come?a com uma marca **pkg:part**.


*Figura 12. Recolher e expandir partes do pacote para facilitar a edi??o no Visual Studio 2015*

![Trecho de c?digo do Office Open XML de uma parte de pacote.](../images/office15-app-create-wd-app-using-ooxml-fig12.png)

*Figura 13. As partes inclu?das em um pacote de documento b?sico do Office Open XML do Word*

![Trecho de c?digo do Office Open XML de uma parte de pacote.](../images/office15-app-create-wd-app-using-ooxml-fig13.png)

Com toda essa marca??o, voc? poder? se surpreender ao descobrir que os ?nicos elementos realmente necess?rios para inserir o exemplo de texto formatado s?o peda?os da parte .rels e a parte document.xml.


    
> [!NOTE]
> As duas linhas de marca??o acima da marca do pacote (as declara??es de XML para a vers?o e a ID do programa do Office) s?o pressupostas quando voc? usa o tipo de coer??o do Office Open XML, assim, n?o ? preciso inclu?-las. Mantenha-as se voc? quiser abrir a marca??o editada como um documento do Word para test?-la.

V?rios dos outros tipos de conte?do mostrados no in?cio deste t?pico tamb?m exigem partes adicionais (al?m daquelas mostradas na Figura 13), e vamos abord?-los mais adiante neste t?pico. Enquanto isso, como voc? ver? a maioria das partes mostradas na Figura 13 na marca??o de qualquer pacote de documento do Word, aqui est? um resumo r?pido do que cada uma das partes faz e quando voc? precisa delas:



- Dentro da marca de pacote, a primeira parte ? o arquivo .rels, que define as rela??es entre as partes de n?vel superior do pacote (elas normalmente s?o as propriedades do documento, a miniatura, se houver, e o corpo do documento principal). Sempre ? necess?rio algum conte?do nessa parte na marca??o, pois voc? precisa definir a rela??o entre a parte do documento principal (em que o conte?do reside) e o pacote de documento.
    
- A parte document.xml.rels define as rela??es para as partes adicionais necess?rias para a parte document.xml (corpo principal), se houver. 
    

    
   > [!IMPORTANT]
   > Os arquivos .rels no pacote (como .rels de n?vel superior, document.xml.rels e outros que voc? pode ver para tipos espec?ficos de conte?do) s?o uma ferramenta extremamente importante que voc? pode usar como guia para ajud?-lo a editar rapidamente o pacote do Office Open XML. Para saber mais sobre como fazer isso, confira [Criar sua pr?pria marca??o: pr?ticas recomendadas](#creating-your-own-markup-best-practices) mais adiante neste t?pico.



- A parte document.xml ? o conte?do no corpo principal do documento. Voc? precisa de elementos dessa parte, claro, pois ? onde o conte?do aparece. Por?m, voc? n?o precisa de tudo o que v? nessa parte. Examinaremos isso em mais detalhes posteriormente.
    
- Muitas partes s?o automaticamente ignoradas pelos m?todos Set ao se inserir conte?do em um documento usando a coer??o do Office Open XML, assim, voc? pode remov?-las. Isso inclui o arquivo theme1.xml (o tema de formata??o do documento), as partes de propriedades do documento (n?cleo, suplemento e miniatura) e arquivos de configura??es (incluindo settings, webSettings e fontTable).
    
- No exemplo da Figura 1, a formata??o de texto ? aplicada diretamente (ou seja, cada configura??o de fonte e de formata??o de par?grafo ? aplicada individualmente). Contudo, se voc? usar um estilo (por exemplo, se desejar que o texto assuma automaticamente a formata??o do estilo T?tulo 1 no documento de destino) como mostrado anteriormente na Figura 2, precisar? da parte styles.xml, bem como de uma defini??o de relacionamento para ele. Para saber mais, confira a se??o do t?pico [Adicionar objetos que usam partes adicionais do Office Open XML](#adding-objects-that-use-additional-office-open-xml-parts).
    

## <a name="inserting-document-content-at-the-selection"></a>Inserir conte?do de documento na sele??o


Vamos examinar a marca??o m?nima do Office Open XML necess?ria para o exemplo de texto formatado mostrado na Figura 1 e o JavaScript necess?rio para inseri-la na sele??o ativa no documento.


### <a name="simplified-office-open-xml-markup"></a>Marca??o simplificada do Office Open XML

Editamos o exemplo do Office Open XML mostrado aqui, conforme descrito na se??o anterior, para deixar apenas as partes do documento obrigat?rias e somente os elementos necess?rios em cada uma dessas partes. Vamos examinar como editar a marca??o voc? mesmo (e explicar um pouco mais as partes restantes aqui) na pr?xima se??o do t?pico.


```XML
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" >
        <w:body>
          <w:p>
            <w:pPr>
              <w:spacing w:before="360" w:after="0" w:line="480" w:lineRule="auto"/>
              <w:rPr>
                <w:color w:val="70AD47" w:themeColor="accent6"/>
                <w:sz w:val="28"/>
              </w:rPr>
            </w:pPr>
            <w:r>
              <w:rPr>
                <w:color w:val="70AD47" w:themeColor="accent6"/>
                <w:sz w:val="28"/>
              </w:rPr>
              <w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t>
            </w:r>
          </w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>
```


> [!NOTE]
> Se voc? adicionar a marca??o mostrada aqui a um arquivo XML com as marcas de declara??o de XML para vers?o e mso-application na parte superior do arquivo (mostrado na Figura 13), voc? poder? abri-lo no Word como um documento do Word. Ou, sem essas marcas, ainda poder? abri-lo usando **Arquivo > Abrir** no Word. Voc? ver? **Modo de Compatibilidade** na barra de t?tulo no Word 2013, pois removeu as configura??es que avisam ao Word que se trata de um documento da vers?o 2013. Como voc? est? adicionando a marca??o a um documento existente do Word 2013, isso n?o afetar? o conte?do de forma alguma.


### <a name="javascript-for-using-setselecteddataasync"></a>JavaScript para usar setSelectedDataAsync


Ap?s salvar o Office Open XML anterior como um arquivo XML que pode ser acessado por meio de sua solu??o, voc? poder? usar a fun??o a seguir para definir o conte?do de texto formatado no documento usando a coer??o do Office Open XML. 

Nessa fun??o, observe que, exceto pela ?ltima linha, tudo ? usado para acessar a marca??o salva para uso na chamada de m?todo [setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) no fim da fun??o. **setSelectedDataASync** requer apenas que voc? especifique o conte?do a ser inserido e o tipo de coer??o.


> [!NOTE]
> Substitua _yourXMLfilename_ pelo nome e pelo caminho do arquivo XML que voc? salvou na solu??o. Se n?o tiver certeza de onde incluir arquivos XML na solu??o ou como referenci?-los no c?digo, confira o exemplo de c?digo [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) para obter exemplos disso e um exemplo operacional da marca??o e do JavaScript mostrado aqui.




```js
function writeContent() {
    var myOOXMLRequest = new XMLHttpRequest();
    var myXML;
    myOOXMLRequest.open('GET', 'yourXMLfilename', false);
    myOOXMLRequest.send();
    if (myOOXMLRequest.status === 200) {
        myXML = myOOXMLRequest.responseText;
    }
    Office.context.document.setSelectedDataAsync(myXML, { coercionType: 'ooxml' });
}
```


## <a name="creating-your-own-markup-best-practices"></a>Criar sua pr?pria marca??o: pr?ticas recomendadas


Vamos examinar mais detalhadamente a marca??o que deve ser inserida no exemplo de texto formatado anterior.

Para o exemplo, comece simplesmente excluindo todas as partes de documento do pacote, exceto .rels e document.xml. Em seguida, editaremos essas duas partes necess?rias para simplificar tudo ainda mais.


> [!IMPORTANT]
> Use as partes .rels como um mapa para avaliar rapidamente o que est? inclu?do no pacote e determinar quais partes voc? pode excluir completamente (ou seja, as partes n?o relacionadas ou nem referenciadas pelo conte?do). Lembre-se de que todas as partes do documento devem ter uma rela??o definida no pacote e as rela??es aparecem nos arquivos .rels. Assim, voc? deve ver todas elas listadas em .rels, em document.xml.rels ou em um arquivo .rels espec?fico do conte?do.

A marca??o a seguir mostra a parte .rels necess?ria antes da edi??o. Como estamos excluindo o suplemento, partes de propriedade do documento principal e a parte de miniatura, tamb?m precisamos excluir essas rela??es de .rels. Observe que isso deixar? somente a rela??o (com a ID de rela??o "rID1" no exemplo a seguir) para document.xml.




```XML
<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
  <pkg:xmlData>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="docProps/thumbnail.emf"/>
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
    </Relationships>
  </pkg:xmlData>
</pkg:part>
```


> [!IMPORTANT]
> Remova as rela??es (ou seja, a marca **Relationship**) de todas as partes que voc? remover completamente do pacote. Incluir uma parte sem uma rela??o correspondente ou excluir uma parte e deixar sua rela??o no pacote resultar? em um erro.

A marca??o a seguir mostra a parte document.xml, que inclui o conte?do de texto formatado de exemplo antes da edi??o.

```XML
<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document mc:Ignorable="w14 w15 wp14" xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
        <w:body>
          <w:p>
            <w:pPr>
              <w:spacing w:before="360" w:after="0" w:line="480" w:lineRule="auto"/>
              <w:rPr>
                <w:color w:val="70AD47" w:themeColor="accent6"/>
                <w:sz w:val="28"/>
              </w:rPr>
            </w:pPr>
            <w:r>
              <w:rPr>
                <w:color w:val="70AD47" w:themeColor="accent6"/>
                <w:sz w:val="28"/>
              </w:rPr>
              <w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t>
            </w:r>
            <w:bookmarkStart w:id="0" w:name="_GoBack"/>
            <w:bookmarkEnd w:id="0"/>
          </w:p>
          <w:p/>
          <w:sectPr>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
            <w:cols w:space="720"/>
          </w:sectPr>
        </w:body>
      </w:document>
    </pkg:xmlData>
</pkg:part>
```

Como document.xml ? a parte do documento principal em que voc? coloca o conte?do, vamos dar uma olhada r?pida nessa parte. (A Figura 14, exibida ap?s a lista, fornece uma refer?ncia visual para mostrar como parte do conte?do principal e das marcas de formata??o explicadas aqui se relacionam ao que voc? v? em um documento do Word.) 


- A marca de abertura **w:document** inclui v?rias listagens de namespaces (**xmlns**). Muitos desses namespaces referem-se a tipos espec?ficos de conte?do, e voc? s? precisa deles caso sejam relevantes para o conte?do.
    
    O prefixo para as marcas em uma parte do documento remete aos namespaces. Neste exemplo, o ?nico prefixo usado nas marcas em todo o document.xml ? **w:**, portanto o ?nico namespace que precisamos deixar na marca de abertura **w:document** ? **xmlns:w**.
    

> [!TIP]
> Se voc? estiver editando a marca??o no Visual Studio de 2015, ap?s excluir namespaces em qualquer parte, examine todas as marcas dessa parte. Se tiver removido um namespace necess?rio para a marca??o, voc? ver? um pequeno sublinhado ondulado vermelho no prefixo relevante das marcas afetadas. Se remover o namespace **xmlns:mc**, voc? tamb?m dever? remover o atributo **mc:Ignorable** que precede as listagens de namespace.


- Dentro da marca de abertura do corpo, voc? ver? uma marca de par?grafo (**w:p**), que inclui o conte?do para este exemplo.
    
- A marca **w:pPr** inclui propriedades para formata??o de par?grafo aplicada diretamente, como um espa?o antes ou depois do par?grafo, o alinhamento do par?grafo ou os recuos. (A formata??o direta refere-se aos atributos que voc? aplica individualmente ao conte?do, n?o como parte de um estilo.) Essa marca tamb?m inclui formata??o de fonte direta que ? aplicada a todo o par?grafo, em uma marca aninhada **w:rPr** (propriedades de execu??o), que cont?m a cor da fonte e o tamanho definido no exemplo.
    

   > [!NOTE]
   > Talvez voc? perceba que tamanhos de fonte e outras configura??es de formata??o na marca??o do Word do Office Open XML parecem ter o dobro do tamanho real. Isso ocorre porque o espa?amento de par?grafo e linha, bem como algumas propriedades de formata??o de se??o mostradas na marca??o anterior, s?o especificados em twips (um vig?simo de um ponto). Dependendo dos tipos de conte?do com os quais trabalha no Office Open XML, voc? pode ver v?rias unidades de medida adicionais, incluindo Unidades M?tricas em Ingl?s (914.400 EMUs para uma polegada), que s?o usadas para alguns valores de Arte do Office (drawingML) e 100.000 vezes o valor real, que ? usado em drawingML e na marca??o do PowerPoint. O PowerPoint tamb?m expressa alguns valores como 100 vezes o valor real, e o Excel comumente usa os valores reais.


- Em um par?grafo, qualquer conte?do com propriedades semelhantes ? inclu?do em uma execu??o (**w:r**), como ? o caso do texto de exemplo. Sempre que h? uma altera??o no tipo de conte?do ou formata??o, uma nova execu??o ? iniciada. (Ou seja, se apenas uma palavra no texto de exemplo estivesse em negrito, ela seria separada em sua pr?pria execu??o.) Neste exemplo, o conte?do inclui apenas o texto de uma execu??o.
    
    Como a formata??o inclu?da neste exemplo ? a formata??o da fonte (ou seja, a formata??o que pode ser aplicada a apenas um caractere), ela tamb?m aparece nas propriedades para a execu??o individual. 
    
- Observe tamb?m as marcas para o indicador oculto "_GoBack" (**w:bookmarkStart** e **w:bookmarkEnd**), que aparecem nos documentos do Word 2013 por padr?o. Voc? sempre pode excluir as marcas de in?cio e de t?rmino do indicador GoBack da marca??o.
    
- A ?ltima parte do corpo do documento ? a marca **w:sectPr**, ou propriedades de se??o. Essa marca inclui configura??es como margens e orienta??o da p?gina. O conte?do que voc? inserir usando **setSelectedDataAsync** adotar? as propriedades da se??o ativa no documento de destino por padr?o. Portanto, a menos que o conte?do inclua uma quebra de se??o (nesse caso, haver? mais de uma marca **w:sectPr**), voc? pode excluir essa marca.
    

*Figura 14. Como marcas comuns em document.xml est?o relacionadas ao conte?do e ao layout de um documento do Word*

![Elementos do Office Open XML em um documento do Word.](../images/office15-app-create-wd-app-using-ooxml-fig14.png)
    
> [!TIP]
> Na marca??o que voc? criar, talvez haja outro atributo em v?rias marcas que inclui os caracteres **w:rsid**, que voc? n?o v? nos exemplos usados neste t?pico. Esses s?o identificadores de revis?o. Eles s?o usados no Word para o recurso Combinar Documentos e est?o ativados por padr?o. Voc? nunca precisar? deles na marca??o que est? inserindo com o suplemento, e desativ?-los torna a marca??o bem mais limpa. Voc? pode facilmente remover marcas RSID existentes ou desabilitar o recurso (conforme descrito no procedimento a seguir) para que eles n?o sejam adicionados ? marca??o para o novo conte?do.
 
Lembre-se de que se voc? usar os recursos de coautoria no Word (como a capacidade de editar simultaneamente documentos com outras pessoas), voc? deve ativar o recurso novamente quando tiver terminado de gerar a marca??o para seu suplemento.
   
Para desativar atributos RSID no Word para documentos que voc? criar no futuro, fa?a o seguinte: 

1. No Word 2013, escolha a guia **Arquivo** e escolha **Op??es**.
2. Na caixa de di?logo Op??es do Word, escolha **Central de Confiabilidade** e escolha **Configura??es da Central de Confiabilidade**.
3. Na caixa de di?logo Central de Confiabilidade, escolha **Op??es de privacidade** e desative a configura??o **Armazenar n?mero aleat?rio para melhorar a precis?o da combina??o**.

Para remover marcas RSID de um documento existente, tente o seguinte atalho com o documento aberto no Office Open XML:


1. Com o ponto de inser??o no corpo do documento principal, pressione **Ctrl+Home** para ir para a parte superior do documento.
2. No teclado, pressione **Barra de espa?os**, **Delete**, **Barra de espa?os**. Em seguida, salve o documento.

Ap?s remover a maior parte da marca??o do pacote, resta a marca??o m?nima que precisa ser inserida para o exemplo, conforme mostrado na se??o anterior.


## <a name="using-the-same-office-open-xml-structure-for-different-content-types"></a>Usar a mesma estrutura do Office Open XML para diferentes tipos de conte?do


V?rios tipos de conte?do avan?ado exigem somente os componentes .rels e document.xml mostrados no exemplo anterior, incluindo controles de conte?do, formas de desenho e caixas de texto do Office e tabelas (a menos que um estilo seja aplicado ? tabela). De fato, voc? pode reutilizar as mesmas partes de pacote editadas e trocar apenas o conte?do de **body** em document.xml para a marca??o do conte?do.

Para verificar a marca??o do Office Open XML para os exemplos de cada um dos tipos de conte?do mostrados anteriormente nas Figuras 5 a 8, explore o exemplo de c?digo [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) referenciado na se??o Vis?o geral.

Antes de continuarmos, vamos dar uma olhada nas diferen?as relevantes para alguns desses tipos de conte?do e como trocar as partes de que voc? necessita.


### <a name="understanding-drawingml-markup-office-graphics-in-word-what-are-fallbacks"></a>Compreender a marca??o de drawingML (elementos gr?ficos do Office) no Word: O que s?o fallbacks?

Se a marca??o da forma ou da caixa de texto parece muito mais complexa do que o esperado, h? um motivo para isso. Com o lan?amento do Office 2007, houve a introdu??o dos Formatos do Office Open XML e de um novo mecanismo de elementos gr?ficos do Office que o PowerPoint e o Excel adotaram plenamente. Na vers?o 2007, o Word s? incorporou parte desse mecanismo de elementos gr?ficos, adotando o mecanismo de elementos gr?ficos atualizado do Excel, elementos gr?ficos SmartArt e ferramentas de imagem avan?adas. Para formas e caixas de texto, o Word 2007 continua a usar objetos de desenho herdados (VML). Na vers?o 2010, o Word lan?ou etapas adicionais com o mecanismo de elementos gr?ficos para incorporar formas e ferramentas de desenho atualizadas.

Portanto, para dar suporte a formas e caixas de texto em documentos do Word no Formato do Office Open XML quando abertos no Word 2007, as formas (incluindo caixas de texto) exigem marca??o VML de fallback.

Normalmente, como voc? pode ver nos exemplos de forma e caixa de texto inclu?dos no exemplo de c?digo [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML), a marca??o de fallback pode ser removida. O Word 2013 adiciona automaticamente a marca??o de fallback ausente ?s formas quando um documento ? salvo. No entanto, se voc? prefere manter a marca??o de fallback para garantir o suporte a todos os cen?rios de usu?rio, n?o h? problema em mant?-la.

Se houver objetos de desenho agrupados inclu?dos no conte?do, voc? ver? marca??o adicional (e aparentemente repetitiva), mas isso deve ser mantido. Partes da marca??o para formas de desenho s?o duplicadas quando o objeto ? inclu?do em um grupo.


> [!IMPORTANT]
> Ao trabalhar com caixas de texto e formas de desenho, verifique os namespaces cuidadosamente antes de remov?-los de document.xml. (Ou ent?o, se voc? estiver reutilizando marca??o de outro tipo de objeto, adicione novamente quaisquer namespaces necess?rios que tenham sido removidos anteriormente de document.xml.) Uma parte substancial dos namespaces inclu?dos por padr?o em document.xml est? presente devido a requisitos de objeto de desenho.


#### <a name="about-graphic-positioning"></a>Sobre o posicionamento de gr?ficos

Nos exemplos de c?digo [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) e [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML), a caixa de texto e a forma s?o configuradas usando diferentes tipos de configura??es de posicionamento e disposi??o de texto. (Lembre-se tamb?m de que os exemplos de imagem nesses exemplos de c?digo s?o configurados usando formata??o embutida com texto, que posiciona um objeto gr?fico na linha de base do texto.)

A forma nesses exemplos de c?digo ? posicionada em rela??o ?s margens direita e inferior da p?gina. O posicionamento relativo permite fazer a coordena??o mais facilmente com a configura??o de documento desconhecida do usu?rio, pois ela se ajustar? ?s margens do usu?rio e haver? menos risco de causar uma apar?ncia estranha devido ?s configura??es de tamanho do papel, orienta??o ou margem. Para manter as configura??es de posicionamento relativas ao inserir um objeto gr?fico, voc? deve manter a marca de par?grafo (w:p) em que o posicionamento (conhecido no Word como uma ?ncora) ? armazenado. Se inserir o conte?do em uma marca de par?grafo existente em vez de incluir a sua pr?prio, voc? poder? manter a mesma apar?ncia inicial, mas muitos tipos de refer?ncias relativas que habilitam o posicionamento a se ajustar automaticamente ao layout do usu?rio poder?o ser perdidos.


### <a name="working-with-content-controls"></a>Trabalho com controles de conte?do

Os controles de conte?do s?o um recurso importante no Word 2013 que pode aprimorar consideravelmente a capacidade do suplemento para o Word de v?rias maneiras, incluindo permitindo-lhe inserir o conte?do em locais designados no documento, em vez de apenas na sele??o.

No Word, localize os controles de conte?do na guia Desenvolvedor da faixa de op??es, conforme mostrado aqui na Figura 15.


*Figura 15. O grupo Controles na guia Desenvolvedor no Word*

![Grupo de Controles de conte?do na faixa de op??es do Word 2013.](../images/office15-app-create-wd-app-using-ooxml-fig15.png)

Os tipos de controles de conte?do no Word incluem RTF, texto sem formata??o, imagem, galeria de blocos de constru??o, caixa de sele??o, lista suspensa, caixa de combina??o, seletor de data e se??o de repeti??o. 



- Use o comando **Propriedades**, mostrado na Figura 15, para editar o t?tulo do controle e para definir prefer?ncias, como ocultar o cont?iner de controle.
    
- Habilite **Modo de Design** para editar o conte?do de espa?o reservado no controle.
    
Se o suplemento funciona com um modelo do Word, voc? pode incluir controles no modelo para aprimorar o comportamento do conte?do. Voc? tamb?m pode usar uma associa??o de dados XML em um documento do Word para associar controles de conte?do a dados, como propriedades de documento, para preencher facilmente formul?rios ou realizar tarefas semelhantes. (Localize os controles que j? est?o associados a propriedades internas do documento no Word na guia **Inserir** em **Partes R?pidas**.)

Ao usar controles de conte?do com o suplemento, voc? tamb?m pode expandir muito as op??es para o que o suplemento pode fazer usando um tipo diferente de associa??o. Voc? pode associar a um controle de conte?do de dentro do suplemento e, depois, escrever conte?do para a associa??o em vez de para a sele??o ativa.


    
> [!NOTE]
> N?o confunda a associa??o de dados XML no Word com a capacidade de associar a um controle por meio do suplemento. Esses s?o recursos completamente separados. No entanto, voc? pode incluir controles de conte?do nomeados no conte?do que inserir por meio do suplemento usando a coer??o de OOXML e usar c?digo no suplemento para associar a esses controles.

Al?m disso, lembre-se de que associa??o de dados XML e o Office.js podem interagir com partes XML personalizadas no aplicativo. Portanto, ? poss?vel integrar essas poderosas ferramentas. Para saber mais sobre como trabalhar com partes XML personalizadas na API JavaScript para Office, confira a se??o [Recursos adicionais](#see-also) deste t?pico.

O trabalho com associa??es no suplemento do Word ? abordado na pr?xima se??o do t?pico. Primeiro, vamos conferir um exemplo do Office Open XML necess?rio para inserir um controle de conte?do RTF que voc? pode associar usando o suplemento.


    
> [!IMPORTANT]
> Controles RTF s?o o ?nico tipo de controle de conte?do que voc? pode usar para associar a um controle de conte?do de dentro do suplemento.




```XML
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" >
        <w:body>
          <w:p/>
          <w:sdt>
              <w:sdtPr>
                <w:alias w:val="MyContentControlTitle"/>
                <w:id w:val="1382295294"/>
                <w15:appearance w15:val="hidden"/>
                <w:showingPlcHdr/>
              </w:sdtPr>
              <w:sdtContent>
                <w:p>
                  <w:r>
                  <w:t>[This text is inside a content control that has its container hidden. You can bind to a content control to add or interact with content at a specified location in the document.]</w:t>
                </w:r>
                </w:p>
              </w:sdtContent>
            </w:sdt>
          </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
 </pkg:package>
```

Como j? mencionado, os controles de conte?do, como texto formatado, n?o exigem partes de documento adicionais. Portanto, somente editadas vers?es das partes .rels e document.xml s?o inclu?das aqui. 

A marca **w:sdt** que voc? v? no corpo de document.xml representa o controle de conte?do. Se gerar a marca??o do Office Open XML para um controle de conte?do, voc? ver? que v?rios atributos foram removidos do exemplo, incluindo a marca e as propriedades de parte de documento. Somente elementos essenciais (e alguns de pr?ticas recomendadas) foram mantidos, incluindo o seguinte:



- O **alias** ? a propriedade de t?tulo da caixa de di?logo Propriedades de Controle de Conte?do no Word. Essa ? uma propriedade necess?ria (representando o nome do item) se voc? planeja associar ao controle de dentro do suplemento.
    
- A **id** exclusiva ? uma propriedade necess?ria. Se voc? associar ao controle de dentro do suplemento, a ID ser? a propriedade que a vincula??o usa no documento para identificar o controle de conte?do nomeado aplic?vel.
    
- O atributo **appearance** ? usado para ocultar o cont?iner de controle, para gerar uma apar?ncia mais limpa. Esse ? um novo recurso no Word 2013, como voc? pode ver pelo uso do namespace w15. Como essa propriedade ? usada, o namespace w15 ? mantido no in?cio da parte document.xml.
    
- O atributo **showingPlcHdr** ? uma configura??o opcional que define o conte?do padr?o que voc? inclui no controle (texto, neste exemplo) como conte?do de espa?o reservado. Portanto, se o usu?rio clica ou toca na ?rea de controle, todo o conte?do ? selecionado, em vez de se comportar como conte?do edit?vel no qual o usu?rio pode fazer altera??es.
    
- Embora a marca de par?grafo vazia (**w:p /**) que precede a marca **sdt** n?o seja necess?ria para adicionar um controle de conte?do (e adicionar? espa?o vertical acima do controle no documento do Word), ela garante que o controle seja colocado em seu pr?prio par?grafo. Isso pode ser importante, dependendo do tipo e da formata??o do conte?do ser? adicionado ao controle.
    
- Se voc? pretende associar ao controle, o conte?do padr?o para o controle (o que est? dentro da marca **sdtContent**) deve incluir pelo menos um par?grafo completo (como neste exemplo), para que a associa??o aceite o conte?do avan?ado com v?rios par?grafos.
    

    
> [!NOTE]
> O atributo de parte de documento que foi removido desta marca de exemplo **w:sdt** pode aparecer em um controle de conte?do para fazer refer?ncia a uma parte separada no pacote em que as informa??es de conte?do de espa?o reservado podem ser armazenadas (partes localizados em um diret?rio de gloss?rio no pacote do Office Open XML). Embora parte de documento seja o termo usado para partes XML (ou seja, arquivos) dentro de um pacote do Office Open XML, o termo partes de documento, conforme usado na propriedade sdt, refere-se ao mesmo termo no Word que ? usado para descrever alguns tipos de conte?do, incluindo blocos de constru??o e partes r?pidas de propriedade de documento (por exemplo, controles associados a dados XML internos). Se houver partes em um diret?rio de gloss?rio no pacote do Office Open XML, talvez voc? precise mant?-las se o conte?do que estiver inserindo incluir esses recursos. Para um controle de conte?do t?pico que voc? pretende usar para associar do suplemento, elas n?o s?o necess?rias. Lembre-se apenas de que, se voc? excluir as partes de gloss?rio do pacote, tamb?m dever? remover o atributo de parte de documento da marca w:sdt.

A pr?xima se??o abordar? como criar e usar associa??es no suplemento do Word.


## <a name="inserting-content-at-a-designated-location"></a>Inserir conte?do em um local designado


J? vimos como inserir o conte?do na sele??o ativa em um documento do Word. Se associar a um controle de conte?do nomeado no documento, voc? poder? inserir qualquer um dos mesmos tipos de conte?do no controle. 

Ent?o, quando conv?m usar essa abordagem?


- Quando voc? precisa adicionar ou substituir conte?do em locais espec?ficos em um modelo, como para preencher partes do documento de um banco de dados
    
- Quando voc? quer a op??o de substituir o conte?do que est? inserindo na sele??o ativa, como para fornecer op??es de elemento de design ao usu?rio
    
- Quando voc? quer que o usu?rio adicione dados no documento que voc? possa acessar para uso com o suplemento, como para preencher campos no painel de tarefas com base em informa??es que o usu?rio adiciona ao documento
    
Baixe o c?digo de exemplo [Word-Add-in-JavaScript-AddPopulateBindings](https://github.com/OfficeDev/Word-Add-in-JavaScript-AddPopulateBindings), que fornece um exemplo de como inserir e associar a um controle de conte?do e como preencher a associa??o.


### <a name="add-and-bind-to-a-named-content-control"></a>Adicionar e associar a um controle de conte?do nomeado


Ao examinar o JavaScript a seguir, considere estes requisitos:


- Conforme mencionado anteriormente, voc? deve usar um controle de conte?do avan?ado para associar ao controle do suplemento do Word.
    
- O controle de conte?do deve ter um nome (esse ? o campo **T?tulo** na caixa de di?logo Propriedades de Controle de Conte?do, que corresponde ? marca **alias** na marca??o do Office Open XML). Isso ? como o c?digo identifica onde colocar a associa??o.
    
- Voc? pode ter v?rios controles nomeados e associ?-los conforme necess?rio. Use um nome de controle de conte?do, uma ID de controle de conte?do e uma ID de associa??o exclusivos.
    

```js
function addAndBindControl() {
    Office.context.document.bindings.addFromNamedItemAsync("MyContentControlTitle", "text", { id: 'myBinding' }, function (result) {
        if (result.status == "failed") {
            if (result.error.message == "The named item does not exist.")
                var myOOXMLRequest = new XMLHttpRequest();
                var myXML;
                myOOXMLRequest.open('GET', '../../Snippets_BindAndPopulate/ContentControl.xml', false);
                myOOXMLRequest.send();
                if (myOOXMLRequest.status === 200) {
                    myXML = myOOXMLRequest.responseText;
                }
                Office.context.document.setSelectedDataAsync(myXML, { coercionType: 'ooxml' }, function (result) {
                    Office.context.document.bindings.addFromNamedItemAsync("MyContentControlTitle", "text", { id: 'myBinding' });
                });
        }
    });
}
```

O c?digo mostrado aqui realiza as seguintes etapas:


- Tenta associar ao controle de conte?do nomeado, usando [addFromNamedItemAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync). 
    
    Execute esta etapa primeiro se houver uma possibilidade para seu suplemento em que o controle nomeado pode j? existir no documento quando o c?digo for executado. Por exemplo, fa?a isto se o suplemento foi inserido em e salvo com um modelo projetado para funcionar com o suplemento, em que o controle foi colocado anteriormente. Voc? tamb?m precisa fazer isto caso necessite associar a um controle que foi colocado anteriormente pelo suplemento.
    
- O retorno de chamada na primeira chamada ao m?todo **addFromNamedItemAsync** verifica o status do resultado para ver se a associa??o falhou porque o item nomeado n?o existe no documento (ou seja, o controle de conte?do chamado MyContentControlTitle neste exemplo). Nesse caso, o c?digo adiciona o controle no ponto de sele??o ativo (usando **setSelectedDataAsync**) e associa a ele.
    

> [!NOTE]
> Como mencionado anteriormente e mostrado no c?digo anterior, o nome do controle de conte?do ? usado para determinar onde criar a associa??o. No entanto, na marca??o do Office Open XML, o c?digo adiciona a associa??o ao documento usando o nome e o atributo de ID do controle de conte?do.

Ap?s a execu??o de c?digo, se examinar a marca??o do documento no qual o suplemento criou associa??es, voc? ver? duas partes para cada associa??o. Na marca??o do controle de conte?do em que uma associa??o foi adicionada (em document.xml), voc? ver? o atributo **w15:webExtensionLinked/**.

Na parte do documento chamada webExtensions1.xml, voc? ver? uma lista das associa??es que criou. Cada uma delas ? identificada usando a ID de associa??o e o atributo de ID do controle aplic?vel, como o item a seguir, em que o atributo **appref** ? a ID de controle de conte?do: ** **we:binding id="myBinding" type="text" appref="1382295294"/**.


> [!IMPORTANT]
> Voc? deve adicionar a associa??o no momento em que pretende agir sobre ela. N?o inclua a marca??o da associa??o no Office Open XML para inserir o controle de conte?do, pois o processo de inser??o dessa marca??o remover? a associa??o.


### <a name="populate-a-binding"></a>Preencher uma associa??o


O c?digo para gravar conte?do para uma associa??o ? semelhante ao usado para gravar conte?do para uma sele??o.


```js
function populateBinding(filename) {
  var myOOXMLRequest = new XMLHttpRequest();
  var myXML;
  myOOXMLRequest.open('GET', filename, false);
  myOOXMLRequest.send();
  if (myOOXMLRequest.status === 200) {
      myXML = myOOXMLRequest.responseText;
  }
  Office.select("bindings#myBinding").setDataAsync(myXML, { coercionType: 'ooxml' });
}
```

Assim como ocorre com **setSelectedDataAsync**, voc? especifica o conte?do a ser inserido e o tipo de coer??o. O ?nico requisito adicional para gravar em uma associa??o ? identific?-la por ID. Observe como a ID de associa??o usada neste c?digo (bindings#myBinding) corresponde ? ID de associa??o estabelecida (myBinding) quando a associa??o foi criada na fun??o anterior.


> [!NOTE]
> O c?digo anterior ? tudo de que voc? precisar? se estiver preenchendo ou substituindo inicialmente o conte?do em uma associa??o. Quando voc? insere um novo item de conte?do em um local associado, o conte?do existente na associa??o ? substitu?do automaticamente. Confira um exemplo disso no exemplo de c?digo referenciado anteriormente, [Word-Add-in-JavaScript-AddPopulateBindings](https://github.com/OfficeDev/Word-Add-in-JavaScript-AddPopulateBindings), que fornece dois exemplos de conte?do separados que voc? pode intercambiar para preencher a mesma associa??o.


## <a name="adding-objects-that-use-additional-office-open-xml-parts"></a>Adicione objetos que usam partes adicionais do Office Open XML


Muitos tipos de conte?do exigem partes adicionais do documento no pacote do Office Open XML, o que significa que fazem refer?ncia a informa??es em outra parte ou o pr?prio conte?do ? armazenado em uma ou mais partes adicionais e referenciado em document.xml.

Por exemplo, considere a seguinte situa??o:


- O conte?do que usa estilos para formata??o (como o texto com estilo mostrado anteriormente na Figura 2 ou a tabela com estilo mostrada na Figura 9) requer a parte styles.xml.
    
- Imagens (como as mostradas na Figuras 3 e 4) incluem os dados de imagem bin?rios em uma e, ?s vezes, em duas partes adicionais.
    
- Diagramas SmartArt (como o que ? mostrado na Figura 10) exigem v?rias partes adicionais para descrever o layout e o conte?do.
    
- Gr?ficos (como o que ? mostrado na Figura 11) exigem v?rias partes adicionais, incluindo sua pr?pria parte de rela??o (.rels).
    
Voc? pode ver exemplos editados da marca??o para todos esses tipos de conte?do no exemplo de c?digo referenciado anteriormente, [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML). Voc? pode inserir todos esses tipos de conte?do usando o mesmo c?digo JavaScript mostrado anteriormente (e fornecido nos exemplos de c?digo referenciados) para inserir o conte?do na sele??o ativa e gravar conte?do em um local espec?fico usando associa??es.

Antes que voc? explore os exemplos, vamos conferir algumas dicas para trabalhar com cada um desses tipos de conte?do.


> [!IMPORTANT]
> Lembre-se: se mantiver partes adicionais referenciadas em document.xml, voc? precisar? manter document.xml.rels e as defini??es de rela??o das partes aplic?veis que est? mantendo, como styles.xml ou um arquivo de imagem.


### <a name="working-with-styles"></a>Como trabalhar com estilos

A mesma abordagem para edi??o de marca??o que vimos no exemplo anterior com texto formatado diretamente ? aplicada ao se usar estilos de par?grafo ou estilos de tabela para formatar o conte?do. No entanto, a marca??o para trabalhar com estilos de par?grafo ? consideravelmente mais simples. Portanto, esse ? o exemplo descrito aqui.


#### <a name="editing-the-markup-for-content-using-paragraph-styles"></a>Editar a marca??o de conte?do usando estilos de par?grafo

A marca??o a seguir representa o conte?do do corpo para o exemplo de texto com estilo mostrado na Figura 2.


```XML
<w:body>
  <w:p>
    <w:pPr>
      <w:pStyle w:val="Heading1"/>
    </w:pPr>
    <w:r>
      <w:t>This text is formatted using the Heading 1 paragraph style.</w:t>
    </w:r>
  </w:p>
</w:body>
```


> [!NOTE]
> Como voc? pode ver, a marca??o de texto formatado em document.xml ? consideravelmente mais simples quando voc? usa um estilo, pois o estilo cont?m toda a formata??o de par?grafo e fonte que, caso contr?rio, voc? precisa referenciar individualmente. No entanto, conforme explicado anteriormente, talvez voc? queira usar estilos ou formata??o direta para fins diferentes: usar formata??o direta para especificar a apar?ncia do texto independentemente da formata??o no documento do usu?rio; usar um estilo de par?grafo (particularmente um nome de estilo de par?grafo interno, como T?tulo 1, mostrado aqui) para que a formata??o do texto seja automaticamente coordenada com o documento do usu?rio.

O uso de um estilo ? um bom exemplo da import?ncia de ler e entender a marca??o para o conte?do que voc? est? inserindo, pois n?o ? expl?cito que outra parte do documento ? referenciada aqui. Se voc? incluir a defini??o de estilo na marca??o e n?o incluir a parte styles.xml, as informa??es de estilo em document.xml ser?o ignoradas independentemente de esse estilo estar ou n?o em uso no documento do usu?rio.

No entanto, se analisar a parte styles.xml, ver? que apenas uma pequena parte dessa longa marca??o ? necess?ria ao editar a marca??o para uso no suplemento:


- A parte styles.xml inclui v?rios namespaces por padr?o. Se voc? estiver mantendo apenas as informa??es de estilo necess?rias para o conte?do, na maioria dos casos, s? precisar? manter o namespace **xmlns:w**.
    
- O conte?do da marca **w:docDefaults** que fica no in?cio da parte styles ser? ignorado quando a marca??o for inserida por meio do suplemento e pode ser removido.
    
- A maior marca??o em uma parte styles.xml ? para a marca **w:latentStyles** que aparece depois de docDefaults, que fornece informa??es (como atributos de apar?ncia para o painel Estilos e a galeria de Estilos) para todos os estilos dispon?veis. Essas informa??es tamb?m ser?o ignoradas ao se inserir conte?do por meio do suplemento e, assim, podem ser removidas.
    
- Ap?s as informa??es de estilos latentes, voc? v? uma defini??o de cada estilo em uso no documento a partir do qual a marca??o foi gerada. Isso inclui alguns estilos padr?o que est?o em uso quando voc? cria um novo documento e podem n?o ser relevantes ao conte?do. Voc? pode excluir as defini??es de estilos que n?o s?o usadas pelo conte?do.
    

   > [!NOTE]
   > Cada estilo de t?tulo interno tem um estilo Char associado que ? uma vers?o de estilo de caractere do mesmo formato do t?tulo. A menos que tenha aplicado o estilo de t?tulo como um estilo de caractere, voc? pode remov?-lo. Se o estilo for usado como um estilo de caractere, ele aparecer? em document.xml em uma marca de propriedades de execu??o (**w:rPr**) em vez de uma marca de propriedades de par?grafo (**w:pPr**). Isso s? dever? ocorrer se voc? tiver aplicado o estilo apenas a parte de um par?grafo, mas poder? ocorrer inadvertidamente se o estilo tiver sido aplicado de forma incorreta.


- Se estiver usando um estilo interno para o conte?do, voc? n?o precisar? incluir uma defini??o completa. Voc? s? deve incluir o nome do estilo, a ID do estilo e pelo menos um atributo de formata??o para que o Office Open XML com coer??o aplique o estilo ao conte?do durante a inser??o.
    
    No entanto, a pr?tica recomendada ? incluir uma defini??o de estilo completa (mesmo que seja o padr?o para os estilos internos). Se um estilo j? estiver sendo usado no documento de destino, seu conte?do adotar? a defini??o do residente para o estilo, independentemente de voc? incluir no styles.xml. Se o estilo ainda n?o estiver sendo usado no documento de destino, seu conte?do usar? a defini??o de estilo que voc? forneceu na marca??o.
    
Portanto, por exemplo, o ?nico conte?do que ? preciso manter da parte styles.xml para o texto de exemplo mostrado na Figura 2, que ? formatado com o estilo T?tulo 1, ? indicado a seguir. 


> [!NOTE]
> Uma defini??o completa do Word 2013 para o estilo T?tulo 1 foi mantida neste exemplo.




```XML
<pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml">
  <pkg:xmlData>
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" >
      <w:style w:type="paragraph" w:styleId="Heading1">
        <w:name w:val="heading 1"/>
        <w:basedOn w:val="Normal"/>
        <w:next w:val="Normal"/>
        <w:link w:val="Heading1Char"/>
        <w:uiPriority w:val="9"/>
        <w:qFormat/>
        <w:pPr>
          <w:keepNext/>
          <w:keepLines/>
          <w:spacing w:before="240" w:after="0" w:line="259" w:lineRule="auto"/>
          <w:outlineLvl w:val="0"/>
        </w:pPr>
        <w:rPr>
          <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
          <w:color w:val="2E74B5" w:themeColor="accent1" w:themeShade="BF"/>
          <w:sz w:val="32"/>
          <w:szCs w:val="32"/>
        </w:rPr>
      </w:style>
    </w:styles>
  </pkg:xmlData>
</pkg:part>
```


#### <a name="editing-the-markup-for-content-using-table-styles"></a>Como editar a marca??o de conte?do usando estilos de tabela


Quando o conte?do usa um estilo de tabela, voc? precisa da mesma parte relativa de styles.xml conforme descrito para trabalhar com estilos de par?grafo. Ou seja, voc? s? precisa manter as informa??es do estilo que est? usando no conte?do e deve incluir o nome, a ID e pelo menos um atributo de formata??o, mas ? melhor incluir uma defini??o de estilo completa para lidar com todos os cen?rios de usu?rio poss?veis.

No entanto, ao conferir a marca??o da tabela em document.xml e da defini??o de estilo de tabela em styles.xml, voc? v? muito mais marca??o do que ao trabalhar com estilos de par?grafo.


- Em document.xml, a formata??o ? aplicada por c?lula, mesmo que esteja inclu?da em um estilo. O uso de um estilo de tabela n?o reduz o volume de marca??o. A vantagem de usar estilos de tabela para o conte?do ? facilitar a atualiza??o e coordenar facilmente a apar?ncia de v?rias tabelas.
    
- Em styles.xml, voc? ver? uma grande quantidade de marca??o para um ?nico estilo de tabela, pois estilos de tabela incluem v?rios tipos de atributos de formata??o poss?veis para cada uma das v?rias ?reas da tabela, como a tabela inteira, linhas de t?tulo, linhas e colunas em faixas pares e ?mpares (separadamente), a primeira coluna etc. 
    

### <a name="working-with-images"></a>Trabalhar com imagens


A marca??o de uma imagem inclui uma refer?ncia a pelo menos uma parte que inclua os dados bin?rios para descrever a imagem. Para uma imagem complexa, isso pode consistir em centenas de p?ginas de marca??o, e voc? n?o pode edit?-la. Como nunca precisa alterar a(s) parte(s) bin?ria(s), voc? poder? simplesmente recolh?-la(s) se estiver usando um editor estruturado como o Visual Studio, para que ainda possa examinar e editar facilmente o restante do pacote.

Se examinar a marca??o de exemplo da imagem simples mostrada anteriormente na Figura 3, dispon?vel no exemplo de c?digo mencionado anteriormente, [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML), voc? ver? que a marca??o da imagem em document.xml inclui informa??es de tamanho e posi??o, bem como uma refer?ncia de rela??o para a parte que cont?m os dados de imagem bin?rios. Essa refer?ncia ? inclu?da na marca **a:blip**, da seguinte maneira:




```XML
<a:blip r:embed="rId4" cstate="print">
```

Lembre-se de que, como uma refer?ncia de rela??o ? usada explicitamente (**r:embed="rID4"**) e essa parte relacionada ? necess?ria para processar a imagem, se voc? n?o incluir os dados bin?rios no pacote do Office Open XML, receber? um erro. Isso ? diferente do styles.xml, explicado anteriormente, que n?o gerar? um erro se for omitido, pois a rela??o n?o ? referenciada explicitamente e a rela??o ? para uma parte que fornece atributos ao conte?do (formata??o), em vez de fazer parte do pr?prio conte?do.


> [!NOTE]
> Ao examinar a marca??o, observe os namespaces adicionais usados na marca a:blip. Voc? ver? em document.xml que o namespace **xlmns:a** (o namespace drawingML principal) ? colocado dinamicamente no in?cio do uso de refer?ncias de drawingML, em vez de no in?cio da parte document.xml. No entanto, o namespace de rela??es (r) deve ser mantido onde aparece no in?cio de document.xml. Verifique se a marca??o de imagem tem requisitos de namespace adicionais. Lembre-se de que voc? n?o precisa memorizar quais tipos de conte?do exigem quais namespaces. Voc? pode distingui-los facilmente examinando os prefixos das marcas em todo o document.xml.


### <a name="understanding-additional-image-parts-and-formatting"></a>No??es b?sicas sobre partes de imagem e formata??o adicionais


Quando voc? usa alguns efeitos de formata??o do Office na imagem, como para a imagem mostrada na Figura 4, que usa configura??es ajustadas de brilho e contraste (al?m de estilo de imagem), pode ser necess?ria uma segunda parte de dados bin?rios de uma c?pia de formato HD dos dados da imagem. Esse formato HD adicional ? necess?rio para a formata??o que ? considerada um efeito de camadas, e a refer?ncia a ele aparece em document.xml, de forma semelhante ao seguinte:


```XML
<a14:imgLayer r:embed="rId5">
```

Veja a marca??o necess?ria para a imagem formatada mostrada na Figura 4 (que usa efeitos das camadas, entre outras) no exemplo de c?digo [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML).


### <a name="working-with-smartart-diagrams"></a>Trabalhar com diagramas SmartArt


Um diagrama SmartArt tem quatro partes associadas, mas apenas duas s?o sempre necess?rias. Voc? pode examinar um exemplo de marca??o SmartArt no exemplo de c?digo [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML). Primeiro, confira uma breve descri??o de cada uma das partes e por que elas s?o necess?rias ou n?o:


> [!NOTE]
> Se o conte?do incluir mais de um diagrama, eles ser?o numerados consecutivamente, substituindo o 1 nos nomes de arquivo listados aqui.


- layout1.xml: essa parte ? necess?ria. Ela inclui a defini??o de marca??o para a funcionalidade e apar?ncia de layout.
    
- data1.xml: essa parte ? necess?ria. Ela inclui os dados em uso na inst?ncia do diagrama.
    
- drawing1.xml: essa parte nem sempre ? necess?ria, mas, se voc? aplicar formata??o personalizada a elementos na inst?ncia de um diagrama, como formatar diretamente formas individuais, talvez seja necess?rio mant?-la.
    
- colors1.xml: essa parte n?o ? necess?ria. Ela inclui informa??es de estilo de cor, mas as cores do diagrama ser?o coordenadas por padr?o com as cores do tema de formata??o ativo no documento de destino, com base no estilo de cor SmartArt que voc? aplicar da guia de design Ferramentas SmartArt no Word antes de salvar a marca??o do Office Open XML.
    
- quickStyles1.xml: essa parte n?o ? necess?ria. De forma semelhante ? parte de cores, voc? pode remov?-la, pois o diagrama adotar? a defini??o do estilo SmartArt aplicado que est? dispon?vel no documento de destino (ou seja, ele ser? coordenado automaticamente com o tema de formata??o no documento de destino).
    

> [!TIP]
> O arquivo SmartArt layout1.xml ? um bom exemplo de locais em que talvez voc? consiga cortar ainda mais a marca??o, mas talvez n?o valha a pena o tempo extra para fazer isso (porque ? removida uma pequena quantidade de marca??o em rela??o ao pacote inteiro). Se quiser remover todas as linhas de marca??o que puder, voc? poder? excluir a marca **dgm:sampData** e seu conte?do. Esses dados de exemplo definem como a visualiza??o de miniatura do diagrama ser? exibida nas galerias de estilos SmartArt. No entanto, se forem omitidos, dados de exemplo padr?o ser?o usados.

Lembre-se de que a marca??o de um diagrama SmartArt em document.xml cont?m refer?ncias de ID de rela??o para partes de layout, dados, cores e estilos r?pidos. Voc? pode excluir as refer?ncias em document.xml das partes de cores e estilos ao excluir essas partes e suas defini??es de rela??o (e certamente ? uma pr?tica recomendada fazer isso, pois voc? est? excluindo essas rela??es), mas n?o receber? um erro se as mantiver, pois n?o s?o necess?rias para que o diagrama seja inserido em um documento. Localize essas refer?ncias em document.xml na marca **dgm:relIds**. Independentemente de voc? executar esta etapa ou n?o, mantenha as refer?ncias de ID de rela??o para as partes de dados e layout necess?rias.


### <a name="working-with-charts"></a>Trabalhar com gr?ficos


De forma semelhante aos diagramas SmartArt, os gr?ficos cont?m v?rias partes adicionais. No entanto, a configura??o para os gr?ficos ? um pouco diferente do SmartArt, pois um gr?fico tem seu pr?prio arquivo de rela??o. A seguir h? uma descri??o das partes de documento necess?rias e remov?veis para um gr?fico:


> [!NOTE]
> Assim como ocorre com diagramas SmartArt, se o conte?do incluir mais de um gr?fico, eles ser?o numerados consecutivamente, substituindo o 1 nos nomes de arquivos listados aqui.


- Em document.xml.rels, voc? ver? uma refer?ncia ? parte necess?ria que cont?m os dados que descrevem o gr?fico (chart1.xml).
    
- Voc? tamb?m ver? um arquivo de rela??o separada para cada gr?fico no pacote do Office Open XML, como chart1.xml.rels.
    
    H? tr?s arquivos referenciados em chart1.xml.rels, mas apenas um ? obrigat?rio. Eles s?o os dados bin?rios da pasta de trabalho do Excel (obrigat?rio) e as cores e partes do estilo (colors1.xml e styles1.xml), que voc? pode remover.
    
Os gr?ficos que voc? pode criar e editar de forma nativa no Word 2013 s?o gr?ficos do Excel 2013, e seus dados s?o mantidos em uma planilha do Excel que ? inserida como dados bin?rios no pacote do Office Open XML. Assim como as partes de dados bin?rios para imagens, esses dados bin?rios do Excel s?o necess?rios, mas n?o h? nada para editar nessa parte. Portanto, voc? pode simplesmente recolher a parte no editor para evitar ter que rolar manualmente por ela para examinar o restante do pacote do Office Open XML.

No entanto, de forma semelhante ao SmartArt, voc? pode excluir as partes de cores e estilos. Se voc? tiver usado os estilos de gr?fico e de cor dispon?veis para formatar o gr?fico, o gr?fico adotar? a formata??o aplic?vel automaticamente quando for inserido no documento de destino.

Confira a marca??o editada do gr?fico de exemplo mostrado na Figura 11 no exemplo de c?digo [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML).


## <a name="editing-the-office-open-xml-for-use-in-your-task-pane-add-in"></a>Edi??o do Office Open XML para uso no suplemento de painel de tarefas


Voc? j? viu como identificar e editar o conte?do na marca??o. Se a tarefa ainda parecer dif?cil quando voc? examinar o enorme pacote do Office Open XML gerado para o documento, veja a seguir um resumo r?pido das etapas recomendadas para ajud?-lo a editar o pacote rapidamente:


> [!NOTE]
> Lembre-se de que voc? pode usar todas as partes .rels no pacote como um mapa para verificar rapidamente se h? partes do documento que pode remover.


1. Abra o arquivo XML compactado no Visual Studio 2015 e pressione Ctrl+K, Ctrl+D para formatar o arquivo. Em seguida, use os bot?es de recolher/expandir ? esquerda para recolher as partes que voc? sabe que precisa remover. Tamb?m conv?m recolher partes longas de que voc? precisa, mas que sabe que n?o precisar? editar (como os dados bin?rios em base 64 para um arquivo de imagem), tornando a verifica??o visual da marca??o mais r?pida e f?cil.
    
2. H? v?rias partes do pacote de documento que voc? quase sempre pode remover ao preparar a marca??o do Office Open XML para uso no suplemento. Conv?m come?ar removendo-as (bem como suas defini??es de rela??o associadas), o que reduzir? bastante o pacote de imediato. Elas incluem theme1, fontTable, settings, webSettings, thumbnail, os arquivos de propriedades principal e do suplemento e quaisquer partes de painel de tarefas ou webExtension.
    
3. Remova as partes que n?o est?o relacionadas ao conte?do, como notas de rodap?, cabe?alhos ou rodap?s dos quais n?o precisa. Novamente, lembre-se de tamb?m excluir suas rela??es associadas.
    
4. Examine a parte document.xml.rels para ver se arquivos referenciados nessa parte s?o necess?rios para o conte?do, como um arquivo de imagem, a parte styles ou partes de diagramas SmartArt. Exclua as rela??es das partes que o conte?do n?o requer e confirme se tamb?m excluiu a parte associada. Se o conte?do n?o exigir nenhuma das partes de documento referenciadas em document.xml.rels, voc? poder? excluir esse arquivo tamb?m.
    
5. Se o conte?do tiver uma parte .rels adicional (como chart#.xml.rels), examine-o para ver se h? outras partes referenciadas que voc? pode remover (como estilos r?pidos para gr?ficos) e exclua a rela??o desse arquivo e a parte associada.
    
6. Edite document.xml para remover namespaces n?o referenciados na parte, propriedades da se??o se o conte?do n?o incluir uma quebra de se??o e qualquer marca??o que n?o esteja relacionada ao conte?do que voc? deseja inserir. Se voc? est? inserindo formas ou caixas de texto, tamb?m conv?m remover a marca??o de fallback ampla.
    
7. Edite quaisquer pe?as necess?rias adicionais em que voc? sabe que pode remover marca??o substancial sem afetar o conte?do, como a parte styles.
    
Ap?s executar as sete etapas anteriores, voc? provavelmente ter? removido cerca de 90% a 100% da marca??o que pode remover, dependendo do conte?do. Na maioria dos casos, provavelmente esse ? o m?ximo que voc? deseja cortar.

Independentemente de voc? parar por aqui ou optar por se aprofundar ainda mais no conte?do para localizar todas as linhas de marca??o que pode recortar, lembre-se de que pode usar o exemplo de c?digo referenciado anteriormente, [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML), como um bloco de rascunho para testar com rapidez e facilidade a marca??o editada.


> [!TIP]
> Se voc? atualizar um trecho do Office Open XML em uma solu??o existente durante o desenvolvimento, limpe arquivos de Internet tempor?rios antes de executar a solu??o novamente para atualizar o Office Open XML usado pelo c?digo. A marca??o inclu?da na solu??o em arquivos XML ? armazenada em cache no computador. Claro, voc? pode limpar os arquivos de Internet tempor?rios do navegador da Web padr?o. Para acessar as op??es da Internet e excluir essas configura??es de dentro do Visual Studio 2015, no menu **Depurar**, escolha **Op??es e Configura??es**. Em seguida, em **Ambiente**, escolha **Navegador da Web** e **Op??es do Internet Explorer**.


## <a name="creating-an-add-in-for-both-template-and-stand-alone-use"></a>Cria??o de um suplemento para o modelo e para uso aut?nomo


Neste t?pico, voc? viu v?rios exemplos do que pode fazer com o Office Open XML em suplementos. Vimos uma ampla variedade de exemplos de tipo de conte?do avan?ado que voc? pode inserir em documentos usando o tipo de coer??o do Office Open XML, juntamente com os m?todos de JavaScript para inserir o conte?do na sele??o ou em um local espec?fico (associado).

Portanto, o que mais voc? precisa saber se estiver criando o suplemento para uso aut?nomo (ou seja, inserido da Loja ou em um local de servidor propriet?rio) e para uso em um modelo pr?-criado projetado para funcionar com o suplemento? A resposta pode ser que voc? j? sabe tudo o que precisa saber.

A marca??o de determinado tipo de conte?do e os m?todos para inseri-la s?o os mesmos, quer o suplemento seja projetado como aut?nomo, quer seja para uso com um modelo. Se voc? estiver usando modelos projetados para funcionar com o suplemento, verifique se o JavaScript inclui retornos de chamada que levam em conta cen?rios em que o conte?do referenciado pode j? existir no documento (conforme demonstrado no exemplo de associa??o mostrado na se??o [Adicionar uma associa??o a um controle de conte?do nomeado](#add-and-bind-to-a-named-content-control)).

Ao usar modelos com o aplicativo, se o suplemento ser? residente no modelo no momento em que o usu?rio criou o documento ou se o suplemento inserir? um modelo, tamb?m conv?m incorporar outros elementos da API para ajud?-lo a criar uma experi?ncia mais robusta e interativa. Por exemplo, conv?m incluir a identifica??o de dados em uma parte customXML que voc? pode usar para determinar o tipo de modelo para oferecer op??es espec?ficas de modelo para o usu?rio. Para saber mais sobre como trabalhar com XML personalizado em suplementos, confira os recursos adicionais a seguir.


## <a name="see-also"></a>Veja tamb?m

- [API JavaScript para Office ](https://dev.office.com/reference/add-ins/javascript-api-for-office) 
- [Padr?o ECMA-376: Formatos do Office Open XML](http://www.ecma-international.org/publications/standards/Ecma-376.htm) (acesse a refer?ncia de linguagem completa e a documenta??o relacionada do Open XML aqui) 
- [OpenXMLDeveloper.org](http://www.openxmldeveloper.org)
- [Como explorar a API JavaScript para Office: associa??o de dados e partes XML personalizadas](https://msdn.microsoft.com/en-us/magazine/dn166930.aspx)
    
