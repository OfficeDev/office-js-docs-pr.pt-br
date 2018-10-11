# <a name="javascript-api-for-office"></a>API JavaScript para Office

A API JavaScript para Office permite criar aplicativos da web que interagem com os modelos de objeto nos aplicativos host do Office. Seu aplicativo fará referência à biblioteca office.js, que é um carregador de script. A biblioteca office.js carrega os modelos de objeto aplicáveis ao aplicativo Office que executa o suplemento. Você pode usar os seguintes modelos de objeto JavaScript:

- **APIs comuns** - APIs que foram introduzidas com o **Office 2013**. É carregado para **todos os aplicativos host do Office** e conecta o seu aplicativo de suplemento com o aplicativo cliente do Office. O modelo de objeto contém APIs específicas para clientes do Office e APIs que se aplicam a vários aplicativos host clientes do Office. Todo esse conteúdo está debaixo da **API Compartilhada**. 

  **O Outlook** também usa a sintaxe da API comum. Tudo que está sob o alias Office contém objetos que você pode usar para escrever scripts que interagem com conteúdo de documentos, planilhas, apresentações, itens de email e projetos do Office a partir do seus suplementos do Office. Você deve usar a API comum se o seu suplemento é direcionado ao Office 2013 e versões posteriores. Esse modelo de objeto usa retornos de chamada.

- **APIs específicas por host** - APIs introduzidas com o **Office 2016**. Este modelo de objeto fornece objetos fortemente tipados específicos para o host que correspondem aos objetos familiares que você vê ao usar os clientes do Office, e representa o futuro das APIs JavaScript para Office. As APIs de host específico incluem atualmente a API JavaScript para Word e a API JavaScript para Excel.

## <a name="supported-host-applications"></a>Aplicativos hosts suportados

- [Excel](overview/excel-add-ins-reference-overview.md)
- [OneNote](overview/onenote-add-ins-javascript-reference.md)
- [Outlook](requirement-sets/outlook-api-requirement-sets.md)
- [Visio](overview/visio-javascript-reference-overview.md)
- [Word](overview/word-add-ins-reference-overview.md)
- [API compartilhada](requirement-sets/office-add-in-requirement-sets.md)

> [!NOTE] 
> [PowerPoint e Project](requirement-sets/powerpoint-and-project-note.md) suportam suplementos feitos com a API JavaScript. No entanto, no momento não têm APIs de host específicas. Você interage com esses hosts por meio da API compartilhada.

Saiba mais sobre [hosts suportados e outros requisitos](https://docs.microsoft.com/office/dev/add-ins/concepts/requirements-for-running-office-add-ins).

## <a name="open-api-specifications"></a>Especificações abertas da API

À medida que criamos e desenvolvemos novas APIs para suplementos do Office, nós as disponibilizamos em nossa página [Especificações abertas da API](openspec.md) a fim de obter os seus comentários. Descubra quais recursos estão no pipeline e comente sobre nossas especificações de design.

## <a name="see-also"></a>Confira também

- [Referência da API JavaScript para Office](https://docs.microsoft.com/javascript/api/overview/office?view=office-js)