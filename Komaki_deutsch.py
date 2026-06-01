import pymupdf
#----------
import re
import unicodedata
import json
import os
import sys
import csv
#----------
import spacy
#----------
from charsplit import Splitter
#----------
import wn
#----------
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt


from pprint import pprint

"""
cefr leveling: 0: removed , 1:A1 , 2:A2 , 3:B1 , 4: upper B1, 5: probably compound word
"""

class komaki_deutsch:
       
    def __init__(self,file_name,cefrlevel:int):

        self.de = wn.Wordnet("odenet:1.4")
        self.en = wn.Wordnet("oewn:2025")

        self.page_dic = [[[]]]# [page number][sentence number]{word: meaning}

        self.determined_cefr_level = cefrlevel

        self.file_full_name = os.path.basename(file_name)
        self.file_name_only = os.path.splitext(self.file_full_name)[0]
        self.dir = os.path.dirname(file_name)

        cefr_to_letter = {0:"A0",1:"A1",2:"A2",3:"B1",4:"UB1"}
        self.json_file_name = f"{self.file_name_only}_KD_{cefr_to_letter[cefrlevel]}.json"
        self.docx_file_name = f"{self.file_name_only}_KD_{cefr_to_letter[cefrlevel]}.docx"

        self.nlp = spacy.load("de_core_news_md")
        # self.nlp = spacy.load("de_core_news_lg")
        # self.nlp = spacy.load("de_dep_news_trf")

        if file_name.endswith(".pdf"):
            self.book_pages = self.get_pdf_pages()
        else:
            print("please only use a pdf file") #---------------NOTICE-------------------

        self.cefr_dict = {}
        with open(".\sprach-o-mat-main\dictionary_a1a2b1_onlystems.csv","r",encoding='utf-8') as csv_file:
            cefr_file = csv.reader(csv_file)
            for row in cefr_file:
                if len(row) >= 3:
                    # {word : cefr-level}
                    self.cefr_dict[row[2].lower()] = row[1]
        self.protect_labels = ["PER", "GPE", "ORG", "MISC", "NORP", "EVENT", "PRODUCT", "WORK_OF_ART"] # LOC is removed

        
            
            #____Configurations____
        self.footer_font = 7
        self.margin = 20
        self.scale_factor_1 = 7
        self.scale_factor_2 = 8
        self.text_font_size = 15
        # 6/10 of text font size 
        self.dict_font_size = 9
        
    def get_pdf_pages(self):
        name = f"{self.dir}\{self.file_full_name}"
        doc = pymupdf.open(name)
        return doc

    def clean_extracted_text(self, text):
        if not text:
            return ""
        
        # Remove control characters
        text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', text)
        
        # Normalize to composed form (preserves umlauts as single characters)
        text = unicodedata.normalize('NFC', text)
        
        # Keep all printable Latin-1 (0-255) characters
        # This includes ä, ö, ü, ß, and other accented letters
        text = ''.join(char for char in text if ord(char) <= 255 or char.isprintable())
        
        # Clean up whitespace
        text = re.sub(r'\s+', ' ', text)
        
        return text.strip()
    
    def make_dictionary(self):
        for page_num,page in enumerate(self.book_pages):
            while len(self.page_dic) <= page_num:
                self.page_dic.append([])
            # print(self.page_dic) #test
            text = self.clean_extracted_text(page.get_text())
            # print(text)
            doc = self.nlp(text)
            for sent_num, sent in enumerate(doc.sents):
                repeatable_words = ""
                #----------experienting----------
                # print(f"\nSent_number: {sent_num}/ {sent}")
                token_num = -1
                for token in sent:
                    if token.text.lower() in repeatable_words:
                        continue
                    repeatable_words += f"{token.text.lower()} "
                    while len(self.page_dic[page_num]) <= sent_num:
                        self.page_dic[page_num].append([])
                    # print(self.page_dic) #test
                    word_cefr = self.get_CEFR_level(token)
                    if word_cefr >= self.determined_cefr_level:
                        # print(self.get_meaningful_words(token.lemma_))
                        #getting the dic_word ready
                        #CHANGED
                        # dic_word = token.lemma_.lower()
                        # if token.pos_ == "VERB" and self.find_separable_preposition(token):
                        #     dic_word = self.find_separable_preposition(token)+ dic_word
                        #CHANGED
                            # print(dic_word)
                        # now we got the words - compound words doesnt need to be perciesly translated
                        # print(f"sent: {sent.text} / word: {dic_word}")
                        # print(self.get_best_definition(sent.text,str(dic_word)))
                        # if self.get_best_definition(sent.text,str(dic_word)):
                            # print(self.get_best_definition(sent.text,str(dic_word))["word"])
                            # print(self.get_best_definition(sent.text,str(dic_word))["en_meaning"])
                        word_dic = self.get_best_definition(sent.text,token)
                        # print(word_dic)
                        if word_dic:
                            self.page_dic[page_num][sent_num].append(word_dic)
                            token_num += 1
                            # TESTING__________
                            # pprint(self.page_dic[page_num][sent_num][token_num])
                            # s_w = self.page_dic[page_num][sent_num][token_num]["searchable_word"]
                            # w_w = self.page_dic[page_num][sent_num][token_num]["word"]
                            # en_m = self.page_dic[page_num][sent_num][token_num]["en_meaning"]
                            # de_m = self.page_dic[page_num][sent_num][token_num]["de_meaning"]
                            # en_l = self.page_dic[page_num][sent_num][token_num]["en_lemma"]
                            # print(f"word: {w_w} =>Lemma: {en_l} => Meaning: {en_m}")

                            # self.get_meaningful_words("Krankenversicherungskarte")
                            # print(self.get_best_definition(sent.text,token)["word"])
                            # print(self.get_best_definition(sent.text,token)["en_meaning"])
                        #if "get the best definition" func returned None , split the word.
        with open(self.json_file_name , 'w', encoding='utf-8') as f:
            json.dump(
                self.page_dic, 
                f, 
                ensure_ascii=False,  # Keep German umlauts (ä, ö, ü, ß)
                indent=2,            # Pretty print with 2 spaces
                sort_keys=False      # Keep original order
            )

        # pprint(self.page_dic)
        # print("\n\n\n")
        # pprint(self.page_dic[0][0][0])




            #----------experienting----------
    
    def get_german_verb_stem(self,word,lemma):
        """
        word: original word form
        lemma: infinitive from spaCy
        """
        # Remove -en or just -n for -ern/-eln verbs
        # if lemma.endswith('eln') or lemma.endswith('ern'):
        #     stem = lemma[:-1]  # remove only -n
        # I intentionally removed "eln" because of the cefr file
        if lemma.endswith('ern'):
            stem = lemma[:-1]  # remove only -n
        elif lemma.endswith('eln'):
            stem = lemma
        else:
            stem = lemma[:-2]  # remove -en
        
        return stem
    
    def replace_umlauts(self, word):
        """
        Replace German umlauts with non-umlaut equivalents.
        'ß': 'ss'
        ä': 'ae' ==> ä': 'a'
        'ü': 'ue' ==> 'ü': 'u'
        """
        umlaut_map = {
            'ä': 'a',
            'ö': 'oe',
            'ü': 'u',
            'ß': 'ss',
            'Ä': 'A',
            'Ö': 'Oe',
            'Ü': 'U'
        }
        
        for umlaut, replacement in umlaut_map.items():
            word = word.replace(umlaut, replacement)
        
        return word.lower()
    
    def get_CEFR_level(self,token):
        
        if token.ent_type_ in self.protect_labels:
            return 0
        
        lemma = token.lemma_.lower()#----------spacy_lemma----------
        
        cefr_map = {"A1":1,"A2":2,"B1":3}

        zero_PREPOSITIONS = [
        "ab", "an", "auf", "aus", "bei", "bis", "durch", "für", "gegen", "hinter",
        "in", "mit", "nach", "neben", "ohne", "über", "um", "unter", "von", "vor",
        "zu", "zwischen", "und", "--","sein","haben"
        ]
        if lemma in zero_PREPOSITIONS:
            return 0
        
        A2_PREPOSITIONS = [
        "außer", "binnen", "entlang", "gegenüber", "seit", "trotz", "während", "wegen"
        ]
        if lemma in A2_PREPOSITIONS:
            return 2
        
        B_plus_PREPOSITIONS = [
        "abzüglich", "anhand", "anlässlich", "anstatt", "aufgrund", "ausschließlich",
        "bezüglich", "diesseits", "einschließlich", "entgegen", "entsprechend",
        "exklusive", "gemäß", "hinsichtlich", "inklusive", "innerhalb", "jenseits",
        "laut", "mithilfe", "nahe", "oberhalb", "samt", "seitens", "statt",
        "trotz", "unterhalb", "unweit", "verzüglich", "wider", "zuzüglich",
        "zugunsten", "zulasten", "zuliebe"
        ]
        if lemma in B_plus_PREPOSITIONS:
            return 4
        
        if token.pos_ == "PRON":
            return 0

        #_____VERB_____
        if token.pos_ == "VERB":

            verb_stem = self.get_german_verb_stem(token,token.lemma_)
            verb_stem_without_umlaut = self.replace_umlauts(verb_stem)
            lemma = verb_stem_without_umlaut #----------cefr_verb_lemma----------
            # print(f"(stem : {verb_stem} )")
            # if verb_stem != verb_stem_without_umlaut:
            #     print(f"(umlaut-less : {verb_stem_without_umlaut} )")
            if self.find_separable_preposition(token):
                svp = self.find_separable_preposition(token)
                full_verb = svp + lemma
                lemma = full_verb
                # print(full_verb)
        #_____VERB_____
        if lemma not in self.cefr_dict:
            def cut_suffix(word):
                """cuts the suffixes"""
                return next((word[:-len(s)] for s in sorted(["keit","lein","ismus","haft","bar","ell","ens","en","er","e","ern"], key=len, reverse=True) if word.endswith(s)), word)
            lemma = cut_suffix(lemma)
            lemma = self.replace_umlauts(lemma)
            # print (f"&%&%&%&&%&% {lemma} %$%$%%$%$%$%")
        
        cefr_level = 5
        if lemma in self.cefr_dict:
            cefr_level = cefr_map[self.cefr_dict[lemma]]
        # print(f"\n word: {token.lemma_} / CEFR_lemma: {lemma} / CEFR_level: {cefr_level} / entity: {token.ent_type_} / POS: {token.pos_}")
        return cefr_level
        # else:
        #     print(f"\n\n word: {token.lemma_} / CEFR_lemma: {lemma} / CEFR_level: {cefr_level}")
        #     return None
    
    def find_separable_preposition(self,token):

        if token.pos_ == "VERB":
            for child in token.children:
                if child.dep_ == "svp": 
                    return child.text
        return None
    
    def word_splitter(self,word):
        splitter = Splitter()
        splitted_words = splitter.split_compound(word)[0]
        return splitted_words
    
    def get_best_definition (self, sentence: str, target_word):

        SPACY_TO_ODENET_POS = {
        "NOUN": "n",
        "PROPN": "n",
        "VERB": "v",
        "AUX": "v",
        "ADJ": "a",
        "ADV": "a",
        }
        if target_word.pos_ in SPACY_TO_ODENET_POS:
            token_pos = SPACY_TO_ODENET_POS[target_word.pos_]
        else:
            token_pos = target_word.pos_
        

        searchable_word = target_word.text
        # Get all synsets for the word
        dic_word = target_word.lemma_.lower()
        if target_word.pos_ == "VERB" and self.find_separable_preposition(target_word):
            dic_word = self.find_separable_preposition(target_word)+ dic_word
        target_word = dic_word
        synsets = self.de.synsets(target_word)

        context_doc = self.nlp(sentence)

        if not synsets:
            if len(self.get_meaningful_words(target_word)) > 1:
                # print(len(self.get_meaningful_words(target_word)))
                # if self.get_meaningful_words(target_word)
                # #__________COMPOUND WORDS__________
                mw_meaning = ""
                meaningful_words = self.get_meaningful_words(target_word)
                for meaningful_word in meaningful_words:
                    best_synset = None
                    best_score = -1
                    best_definition = ""
                    mw_en_lemma = ""
                    mw_en_definition = ""

                    mw_synsets = self.de.synsets(meaningful_word)
                    for mw_synset in mw_synsets:
                        mw_definition = mw_synset.definition()
                        if mw_definition == None:
                            continue
                        mw_doc = self.nlp(mw_definition)
                        similarity = context_doc.similarity(mw_doc)
                        if best_score < similarity:
                            best_score = similarity
                            best_synset = mw_synset
                            best_definition = mw_definition
                        
                    if best_synset:
                        en_synsets = best_synset.translate("oewn:2025")
                        if en_synsets:
                            en_synset = en_synsets[0]
                            mw_en_definition = en_synset.definition()
                            english_lemmas = [lemma for lemma in en_synset.lemmas()]
                            if english_lemmas:
                                mw_en_lemma = english_lemmas[0]
                            else:
                                mw_en_lemma = ""
                        else:
                            mw_en_definition = ""
                            mw_en_lemma = ""
                        mw_meaning += f"\n\t{meaningful_word}:{mw_en_lemma}/{mw_en_definition}"
                en_definition = mw_meaning
                en_lemma = ""
                # print(en_definition)
                return {
                    "word": target_word,
                    "sentence": sentence,
                    "de_meaning": best_definition,
                    "en_meaning": en_definition,
                    "searchable_word":searchable_word,
                    "en_lemma":en_lemma
                }
                #__________COMPOUND WORDS__________   
            else:
                return None 
        
        # Get context words (exclude target word)
        # context_words = [token.text.lower() for token in doc if token.text.lower() != target_word.lower()]
        # context_text = " ".join(context_words)
        # context_doc = self.nlp(context_text)
        
        # Score each synset based on definition similarity
        best_synset = None
        best_score = -1
        best_definition = ""
        
        for synset in synsets:
            odenet_pos = synset.pos
            if token_pos == "n" or token_pos == "a" or token_pos == "v":
                print(f"{searchable_word} {odenet_pos} / {token_pos}")
                if synset.pos != token_pos:
                    print("Ignored!")
                    continue
            definition = synset.definition()
            if definition == None:
                continue
            def_doc = self.nlp(definition)
            
            # Calculate similarity between context and definition
            if context_doc.has_vector and def_doc.has_vector:
                similarity = context_doc.similarity(def_doc)
                
                if similarity > best_score:
                    best_score = similarity
                    best_synset = synset
                    best_definition = definition
        
        if best_synset:
            en_synsets = best_synset.translate("oewn:2025")
            if en_synsets:
                en_synset = en_synsets[0]
                en_definition = en_synset.definition()
                english_lemmas = [lemma for lemma in en_synset.lemmas()]
                if english_lemmas:
                    en_lemma = english_lemmas[0]
                else:
                    en_lemma = ""
            else:
                en_definition = None
                en_lemma = ""

            

            return {
                "word": target_word,
                "sentence": sentence,
                "de_meaning": best_definition,
                "en_meaning": en_definition,
                "searchable_word":searchable_word,
                "en_lemma":en_lemma
            }
        return None
    
    def word_splitter(self, word):
        """Split a German compound word."""
        splitter = Splitter()
        splitted_words = splitter.split_compound(word)[0]
        return splitted_words  # Returns (score, word1, word2)

    def get_meaningful_words(self, word):
        # Use a queue/list with index instead of modifying during iteration
        words_to_process = [word]
        processed_words = set()  # Track processed words to avoid infinite loops
        meaningful_words = []    # Store words that HAVE meanings
        index = 0
        
        def get_lemma(w):
            doc = self.nlp(w)
            return "".join([token.lemma_ for token in doc]) if len(doc) > 0 else w

        while index < len(words_to_process):
            current_word = words_to_process[index]
            index += 1
            
            # Skip if already processed
            if current_word in processed_words:
                continue
            
            processed_words.add(current_word)
            
            # Search for synsets
            g_synsets = self.de.synsets(current_word)
            
            if g_synsets:
                # Word has meaning - add it to results
                meaningful_words.append(current_word)
            else:
                # Word not found - try to split it
                split_result = self.word_splitter(current_word)
                
                # split_result returns a tuple: (score, word1, word2)
                if split_result and len(split_result) >= 3:
                    word1 = get_lemma(split_result[1])
                    word2 = get_lemma(split_result[2])
                    
                    # Add the split words to process
                    if word1 and word1 not in processed_words:
                        words_to_process.append(word1)
                    if word2 and word2 not in processed_words:
                        words_to_process.append(word2)


        
        return meaningful_words

    #__________MAKING PROCESSED FILES__________

    def make_line_by_line(self,pages_dic):

        #_____________________bolding words in paragraph_______________________
        def add_bold_words(paragraph, sentence, bold_words):
            words = sentence.split()
            for word in words:
                run = paragraph.add_run(word + " ")
                run.font.size = Pt(self.text_font_size)
                #Change#1
                # print(word)
                # print(bold_words)
                if word.strip('.,!?\"\':;*()+=\\/<>}{][') in bold_words:
                    # print(f'bold word: {word}')
                    run.bold = True
        #_____________________bolding words in paragraph_______________________

        def set_paragraph_background(paragraph, color="D3D3D3"):
            """Set background shading for a paragraph (color is hex string, e.g., 'FFFF00')"""
            p = paragraph._p  # get the lxml element of the paragraph
            pPr = p.get_or_add_pPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), color)
            pPr.append(shd)

        docx = Document()
        for page_num, page in enumerate(self.book_pages):
            page_dic = pages_dic[page_num]
            text = page.get_text()
            #_____FIXING_____
            text = self.clean_extracted_text(text)
            #_____FIXING_____
            doc = self.nlp(text)
            sents = ""
            for sent_num, sent in enumerate(doc.sents):
                sent_text = sent.text.replace("\n"," ") #__________WHY?!__________
                # docx.add_paragraph(sent_text)
                sents += sent_text+" " #__________until here we have the sentence
                bold_words = []
                sent_dic = page_dic[sent_num]
                for word_d in sent_dic:
                    bold_words.append(word_d["searchable_word"])


                if sent_dic:
                    main_para = docx.add_paragraph()
                    add_bold_words(main_para,sents,bold_words) #sent_dic = bold words =>> array of bold words [X]
                    sents =""
                    for wrd in sent_dic:
                        para = docx.add_paragraph()
                        # making the difficult words bold
                        # run_word = para.add_run(f"{wrd}").bold = True
                        
                        run_word = para.add_run(wrd["word"]) #=> word
                        run_word.font.size = Pt(self.dict_font_size)
                        run_word.bold = True
                        
                        wrd_lemma = wrd["en_lemma"]
                        wrd_meaning = wrd["en_meaning"]
                        run_def = para.add_run(f" : {wrd_lemma}/ {wrd_meaning}") #==> f"{en_lemma}\{en_meaning}"
                        run_def.font.size = Pt(self.dict_font_size)

                        set_paragraph_background(para)
            # main_para = docx.add_paragraph(sents)
            main_para = docx.add_paragraph()
            run_main = main_para.add_run(sents)
            run_main.font.size = Pt(self.text_font_size)
        # print("ta inja ok")
        # docx.save("kir.docx")
        # print(f"{self.dir}/{self.docx_file_name}")
        docx.save(f"{self.dir}\{self.docx_file_name}")

    
    #__________MAKING PROCESSED FILES__________
            

            












if __name__ == "__main__":
    # file_name = r"D:\Python Projects\Komaki_deutsch\few_sents.pdf"
    # file_name = "sentence.pdf"
    # file_name = r"D:\Python Projects\Komaki_deutsch\example.pdf"
    file_name = r"D:\Python Projects\Komaki_deutsch\german_example.pdf"
    # file_name = r"D:\Python Projects\Komaki_deutsch\test.pdf"
    book = komaki_deutsch(file_name,2)
    book.make_dictionary()
    print("making dictionary done!")
    with open(f"{book.json_file_name}","r",encoding='utf-8') as f:
        dictionary = json.load(f)
    book.make_line_by_line(dictionary)
    print("done")

