import type { Catalog } from '../message.js'

export const PRIVACY = {
  'privacy.title': { pt: 'Inspecionar documento', en: 'Inspect document' },
  'privacy.intro': {
    pt: 'Além do conteúdo, o arquivo pode revelar quem o fez e onde. Marque o que remover.',
    en: 'Besides its content, the file may reveal who made it and where. Check what to remove.',
  },
  'privacy.loading': { pt: 'Inspecionando o arquivo…', en: 'Inspecting the file…' },
  'privacy.needsSave': {
    pt: 'O inspetor lê o arquivo como está salvo. Salve antes, para que ele veja o que você acabou de mudar.',
    en: 'The inspector reads the file as saved. Save first, so it sees what you just changed.',
  },
  'privacy.saveAndInspect': { pt: 'Salvar e inspecionar', en: 'Save and inspect' },
  'privacy.nothingFound': {
    pt: 'Nada encontrado: o arquivo não tem nomes, pastas nem metadados de imagem que o Librevia saiba remover.',
    en: 'Nothing found: the file has no names, folders or image metadata Librevia knows how to remove.',
  },
  'privacy.names': { pt: 'Nomes de pessoas e da empresa', en: 'Names of people and the company' },
  'privacy.namesDetail': {
    pt: 'Autor, último a salvar, gerente, empresa e autores de comentários e alterações, que passam a "{anonymous}": {names}',
    en: 'Author, last saved by, manager, company and the authors of comments and changes, who become "{anonymous}": {names}',
  },
  'privacy.paths': { pt: 'Pastas do computador', en: 'Folders on the computer' },
  'privacy.pathsDetail': {
    pt: 'O modelo anexado ao documento ou a última pasta onde o Excel salvou: {paths}',
    en: 'The template attached to the document or the last folder Excel saved in: {paths}',
  },
  'privacy.customProperties': { pt: 'Propriedades personalizadas', en: 'Custom properties' },
  'privacy.customPropertiesDetail': {
    pt: 'Campos criados por pessoas ou sistemas da empresa: {names}',
    en: 'Fields created by people or company systems: {names}',
  },
  'privacy.thumbnail': { pt: 'Miniatura da primeira página', en: 'Thumbnail of the first page' },
  'privacy.thumbnailDetail': {
    pt: 'A imagem que os gerenciadores de arquivos mostram sem abrir o documento.',
    en: 'The picture file managers show without opening the document.',
  },
  'privacy.imageMetadata': { pt: 'Metadados das imagens', en: 'Image metadata' },
  'privacy.imageMetadataDetail': {
    pt: {
      one: '{count} imagem guarda dados da câmera ou do programa, às vezes com o local GPS. A imagem e a orientação ficam.',
      other:
        '{count} imagens guardam dados da câmera ou do programa, às vezes com o local GPS. As imagens e a orientação ficam.',
    },
    en: {
      one: '{count} image keeps camera or program data, sometimes with the GPS location. The image and its orientation stay.',
      other:
        '{count} images keep camera or program data, sometimes with the GPS location. The images and their orientation stay.',
    },
  },
  'privacy.alreadyAnonymous': {
    pt: 'Este arquivo já pede para não guardar nomes: comentários e alterações novos saem como "{anonymous}".',
    en: 'This file already asks not to keep names: new comments and changes come out as "{anonymous}".',
  },
  'privacy.remove': { pt: 'Remover', en: 'Remove' },
  'privacy.removed': {
    pt: 'Removido. Salve o arquivo para gravar a versão sem esses dados; fechar sem salvar mantém o arquivo como estava.',
    en: 'Removed. Save the file to write the version without this data; closing without saving keeps the file as it was.',
  },
  'privacy.readOnly': {
    pt: 'O arquivo está só para leitura. Libere a edição para remover dados.',
    en: 'The file is read only. Allow editing to remove data.',
  },
} satisfies Catalog
