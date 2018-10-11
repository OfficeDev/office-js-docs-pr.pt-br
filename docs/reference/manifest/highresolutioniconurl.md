# <a name="highresolutioniconurl-element"></a>Elemento HighResolutionIconUrl

Especifica a URL da imagem usada para representar seu suplemento do Office no UX e no Office Store de inserção em telas de alto DPI.

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a>Pode conter

[Substituição](override.md)

## <a name="attributes"></a>Atributos

|**Atributo**|**Tipo**|**Obrigatório**|**Descrição**|
|:-----|:-----|:-----|:-----|
|DefaultValue|string (URL)|obrigatório|Especifica o valor padrão para essa configuração, expresso para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).|

## <a name="remarks"></a>Comentários

Para um suplemento de email, o ícone é exibido na interface de usuário **Arquivo**  >  **Gerenciar suplementos**. Para um suplemento de conteúdo ou de painel de tarefas, o ícone é exibido na interface de usuário **Inserir**  >  **Suplementos**.

A imagem deve estar em um dos seguintes formatos de arquivo em uma resolução recomendada de 64 x 64 pixels: GIF, JPG, PNG, EXIF, BMP ou TIFF. Para obter mais informações, consulte a seção _Criar uma identidade visual consistente para seu aplicativo_ em [Criar listagens eficazes na AppSource e no Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings).
