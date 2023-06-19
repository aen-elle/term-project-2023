import codecs, os
import docx2txt
import pymystem3
from nltk import sent_tokenize
import gensim
from re import sub

morph = pymystem3.Mystem(entire_input=False)

PATH = None #для использования необходимо заменить переменную на путь к директории, в которой содержатся текстовые файлы


def get_lemma_POS(item): #лемматизация и частеречная разметка слова 
    if item['text'] == 'BREAK':
        return '\n'  
    elif len(item) == 2 and item['analysis']:
        lemma = item['analysis'][0]['lex']
        grammar = item['analysis'][0]['gr'].split('=')[0]
        grammar = grammar.split(',')[0]
        return(f'{lemma}_{grammar}')
    else:
        return ''

def normalize_text(read_docx_file_path, write_file_path) #удаление из текста лишних символов для упрощения морфологической обработки
    text = docx2txt.process(read_docx_file_path)
    norm_text = sub('\d', '', text)
    norm_text = sub('\n+|\s+', ' ', norm_text)
    norm_text = sub('ї', 'и', norm_text)
    with open(write_file_path, 'w', encoding='utf-8') as f:
        f.write(norm_text)



def morph_process_txt(source_file, dest_file): #лемматизация и частеречная разметка текста
    with open(source_file, 'r', encoding='utf-8') as f:
        text = f.read()
        sentences = sent_tokenize(text, language='russian')
        tokenized_text = ' BREAK '.join(sentences)

    with open(dest_file, 'a', encoding='utf-8') as g:
            morphed = morph.analyze(tokenized_text)
            result = [get_lemma_POS(item) for item in morphed]
            to_write = ' '.join(result)
            g.write(to_write)

    return len(result)

                    


def create_corpus(source_path, dest_file): #создание лемматизированного корпуса текстов
    if os.path.exists(dest_file):
        print('В данном файле уже сохранены данные корпуса. Пожалуйста, выберите другой путь.')
        return
    num_of_tokens = 0
    file_names = os.listdir(source_path)
    for file in file_names:
        txt_name = f'{source_path}/{file[:-4]}'
        if file.endswith('docx'): 
            normalize_text(f'{source_path}/{file}', f'{txt_name}.txt')
            tokens = morph_process_txt(f'{txt_name}.txt', dest_file)
            num_of_tokens = num_of_tokens + tokens
    print(f'You have a brand new corpus! There are {num_of_tokens} tokens in it!')

        

def get_model(source_path, model_name): #создание векторной модели текста на основе корпуса
    model = gensim.models.Word2Vec(corpus_file=source_path,vector_size=8000,window=10,min_count=5,sg=1,epochs=10)
    model.save(f'{model_name}.model')


def get_lsg(model_path, word, num_of_words): #создание лексико-семантической группы для заданного слова
    model = gensim.models.Word2Vec.load(model_path)
    lsg = []
    source_POS = word.split('_')[1]
    try:
        sims = model.wv.most_similar(word, topn = num_of_words)
        for sim in sims:
            POS = sim[0].split('_')[1]
            if POS == source_POS:
                lsg.append(sim)
        print(lsg)
    except KeyError:
        print('Такого слова нет в словаре.')


model_paths = ['models/domovodstvo.model',
               'models/klinker.model',
               'models/geography.model',
               'models/joined.model',
               'models/all_corpus.model']


for i in model_paths: #текстовое слово: можно заменить на любое другое, однако необходимо включить частеречный тэг из разметки Mystem
    get_lsg(i, 'ехать_V', 40)
