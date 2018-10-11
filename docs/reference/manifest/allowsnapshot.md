# <a name="allowsnapshot-element"></a>Elemento AllowSnapshot

Especifica se uma imagem instantânea do seu suplemento de conteúdo é gravada com o documento host.

**Tipo de suplemento:** Conteúdo

## <a name="syntax"></a>Sintaxe

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a>Contido em

[OfficeApp](officeapp.md)

## <a name="remarks"></a>Comentários

 > [!IMPORTANT]
 > **AllowSnapshot** é `true` por padrão. Isso cria uma imagem do suplemento visível para os usuários que abrirem o documento em uma versão do aplicativo host que não oferece suporte a suplementos do Office, ou fornece uma imagem estática do suplemento se o aplicativo host não se conectar ao servidor que hospeda o suplemento. No entanto, isso também significa que informações potencialmente confidenciais exibidas no suplemento podem ser acessadas diretamente a partir do documento que hospeda o suplemento.

