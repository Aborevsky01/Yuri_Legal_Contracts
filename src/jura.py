import torch
from transformers.generation.utils import GenerationConfig
from transformers import AutoModelForCausalLM, AutoTokenizer, AutoConfig
import numpy as np
import os
import sys
import warnings
warnings.filterwarnings('ignore')
import traceback
import json
import asyncio
from tqdm.notebook import tqdm
import nest_asyncio
import json_repair

#from doc_processing import json_to_doc
DEFAULT_SYSTEM_PROMPT = "answer only in json format"

class Jura():

    default_llm_params = {
        'pretrained_model_name_or_path' : "Vikhrmodels/Vikhr-Nemo-12B-Instruct-R-21-09-24", #"IlyaGusev/saiga_llama3_8b", 
        'load_in_8bit' : True,
        'torch_dtype'  : torch.bfloat16,
        'device_map'   : "auto",
    }
    
    def __init__(self, classifier_params=None, law_params=None, ner_params=None):
        self.classification = {
                                'params'      : self.default_llm_params if classifier_params is None else classifier_params,
                                'kw_map_path' : 'docs/km_classification.txt',
                                'tokenizer'   : None,
                                'model'       : None,
                                'gen_config'  : None
        }
        
        self.law = {
                                'params'      : self.default_llm_params if law_params is None else law_params,
                                'kw_map_path' : {
                                        'dcp'     : 'docs/kw_maps_law/km_dcp.txt',
                                        'uslugi'  : 'docs/kw_maps_law/km_uslugi.txt',
                                        'zaym'    : 'docs/kw_maps_law/km_zaym.txt'
                                },
                                'tokenizer'   : None,
                                'model'       : None,
                                'gen_config'  : None
        }
        
        self.ner = {
                                'params'      : self.default_llm_params if ner_params is None else ner_params,
                                'kw_map_path' : {
                                        'dcp'     : 'docs/kw_maps_ner/km_dcp.txt',
                                        'uslugi'  : 'docs/kw_maps_ner/km_uslugi.txt',
                                        'zaym'    : 'docs/kw_maps_ner/km_zaym.txt'
                                },
                                'tokenizer'   : None,
                                'model'       : None,
                                'gen_config'  : None
        }
        self.document_class, self.law_check, self.json_result = None, None, None
        self.max_attempts    = 0
        self.uploaded_models = {}
        
        
    def setup(self, module):
        phase = eval('self.' + module)
        
        if phase['params']['pretrained_model_name_or_path'] in self.uploaded_models.keys():
            phase['model'] = self.uploaded_models[phase['params']['pretrained_model_name_or_path']]
        else:
            phase['model'] = AutoModelForCausalLM.from_pretrained(**phase['params'],  local_files_only = True).eval()
            self.uploaded_models[phase['params']['pretrained_model_name_or_path']] = phase['model']
            
        phase['tokenizer']  = AutoTokenizer.from_pretrained(phase['params']['pretrained_model_name_or_path'], local_files_only = True)
        
        try:
            phase['gen_config'] = GenerationConfig.from_pretrained(phase['params']['pretrained_model_name_or_path'])
            phase['gen_config'].temperature = 0.05
            phase['gen_config'].max_tokens  = 5000
            phase['gen_config'].max_length  = 5000
        
        except Exception as e:
            print('Oh!')
            pass
            
    
    async def async_generate(self, llm, prompt, generation_config=None):
        resp = llm.generate(**prompt, generation_config=generation_config)[0] if True else llm.generate(**prompt)[0]
        #return json.loads("{ " + self.extract_substring(resp[len(prompt['input_ids'][0]):]).replace("\n", "").replace("[]", '""').strip() + "}")
        return resp[len(prompt['input_ids'][0]):]


    async def generate_concurrently(self, llm, prompts, generation_config=None):
        tasks  = [self.async_generate(llm, prompt, generation_config) for prompt in tqdm(prompts)]
        result =  await asyncio.gather(*tasks)
        return result

    def run_model(self, query, phase): 
        file_name = phase['kw_map_path']
        if isinstance(file_name, dict): file_name = file_name[self.document_class]
        with open(file_name, 'r') as file: kw_map = file.read()
        
            
        if kw_map.find('БЛОК') != -1:
            kw_map = kw_map.split('БЛОК')
            instructions = list(map(lambda x : kw_map[0] + x + kw_map[-1], kw_map[1:-1]))
            
            prompts = [phase['tokenizer'].apply_chat_template([
                            {"role": "system", "content": DEFAULT_SYSTEM_PROMPT}, 
                            {"role": "user",  "content": kw + query}], tokenize=False, add_generation_prompt=True) 
                            for kw in instructions
            ]
            
            data = [phase['tokenizer'](prompt, return_tensors="pt", add_special_tokens=False) for prompt in prompts]
            data = [{k: v.to(phase['model'].device) for k, v in x.items()} for x in data]

            nest_asyncio.apply()
            loop       = asyncio.get_event_loop()
            output_ids = loop.run_until_complete(self.generate_concurrently(phase['model'], data, phase['gen_config']))
            
            for i in range(len(output_ids)):
                print(phase['tokenizer'].decode(output_ids[i], skip_special_tokens=True).strip())
                print('----')
            
            output     = [
                json_repair.loads("{ " + self.extract_substring(phase['tokenizer'].decode(out, 
                                        skip_special_tokens=True).strip()).replace("\n", "").replace("[]", '""').strip() + "}")
                for out in tqdm(output_ids)
            ]
            
            for i in range(1, len(output)): 
                output[0].update(output[i])
            output = output[0]

        else:
            prompt = phase['tokenizer'].apply_chat_template([{"role": "system", "content": DEFAULT_SYSTEM_PROMPT}, 
                                                             {"role": "user",  "content": kw_map + query}], 
                                                            tokenize=False, add_generation_prompt=True)
            data = phase['tokenizer'](prompt, return_tensors="pt", add_special_tokens=False)
            data = {k: v.to(phase['model'].device) for k, v in data.items()}

            output_ids = phase['model'].generate(**data, generation_config=phase['gen_config'])[0] if phase['gen_config'] else phase['model'].generate(**prompt)[0]
            output_ids = output_ids[len(data["input_ids"][0]):]
            output = phase['tokenizer'].decode(output_ids, skip_special_tokens=True).strip()
            output = json_repair.loads("{ " + self.extract_substring(output).replace("\n", "").replace("[]", '""').strip() + "}")
            print(output)
       
        return output
        
        
    def launch(self, query, path_to_file=""):
        attempt = 0
        while attempt <= self.max_attempts:
            try: # не делай повторные запуски пройденных этапов
                #=============
                
                if self.document_class is None:
                    if self.classification['model'] is None: self.setup('classification')
                    classifier_response = self.run_model(query, self.classification)
                    if int(classifier_response['Класс']) not in [1, 2, 5, 6]:
                        print('Данный класс договора временно не генерируется. Наиболее качественные результаты в данный момент для услуг, займа и купли-продажи.')
                        return None
                    self.document_class = self.detect_class(classifier_response)

                    if self.document_class is None:
                        cls_attempts = 1
                        while cls_attempts < 3 and self.document_class is None:
                            self.document_class = self.detect_class(classifier_response)
                        if cls_attempts == 3:
                            print('No class identified')
                            return
                print(f'Class identified: {self.document_class}')
                
                #=============
                if self.law_check is None:
                    if self.law['model'] is None: self.setup('law')
                    self.law_check = self.run_model(query, self.law)

                    for key, value in self.law_check.items():
                        if int(value) == 1: # regex
                            print(f'Law {key} was broken')
                            return
                print('Law was not broken')
                
                #=============
                if self.json_result is None:
                    if self.ner['model'] is None: self.setup('ner') # проверка на загрузку одноименной модели
                    self.json_result = self.run_model(query, self.ner) 
                
                    '''
                    for key in list(filter(lambda x: not json_result[x], json_result.keys())):
                        result = self.ask_empty_entity(x)
                        if result is not None:
                            json_result[x] = result
                    '''
                print('Json is ready')
            
                #=============
                path_to_file = json_to_doc(self.json_result, doc_class=self.document_class, path_to_file=path_to_file,
                                          model=self.ner['model'], tokenizer=self.ner['tokenizer'], gen_config=self.ner['gen_config'])
                return path_to_file
                
            except Exception as e:
                attempt += 1
                print(f"Attempt {attempt} failed: {type(e).__name__}: {e}")
                
                exc_type, exc_value, exc_traceback = sys.exc_info()
                print(f"Line Number: {exc_traceback.tb_lineno}")
                print(f"Full Traceback:\n{traceback.format_exc()}")
                
                if attempt > self.max_attempts:
                    print(f"All {self.max_attempts} attempts failed.")
                    return None
                continue
        
            
        
    def extract_substring(self, s: str) -> str:
        start = s.find('{')
        end = s.rfind('}')
        s = s.replace("'", '"')
        
        if start != -1 and end != -1 and start < end:
            return s[start+1:end]
        else:
            return ""
        
    def detect_class(self, cls_res):
        pred  = int(cls_res['Класс'])
        names = {
            1 : 'dcp',
            2 : 'uslugi',
            5 : 'uslugi',
            6 : 'zaym'
        }
    
        return names[pred]
    
