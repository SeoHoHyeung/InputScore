from abc import ABC, abstractmethod
import win32com.client
import time
import threading
from queue import Queue, Empty
import re

class ITTSManager(ABC):
    @abstractmethod
    def speak_name(self, name: str):
        pass

class TTSManager(ITTSManager):
    _instance = None
    _lock = threading.Lock()
    
    def __new__(cls):
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super().__new__(cls)
                    cls._instance._initialized = False
        return cls._instance
    
    def __init__(self):
        if self._initialized:
            return
            
        self._initialized = True
        self.speaker = None
        self.last_speak_time = 0
        self.last_spoken_name = None
        
        # 성능 최적화를 위한 변수들
        self._tts_queue = Queue(maxsize=5)  # TTS 큐 크기 제한
        self._worker_thread = None
        self._is_running = False
        self._name_cache = {}  # 이름 처리 결과 캐싱
        self._english_pattern = re.compile(r'[a-zA-Z]')  # 영어 패턴 컴파일
        
        self.setup()
    
    def setup(self):
        """TTS 엔진을 초기화합니다."""
        try:
            self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
            # 음성 속성 설정
            self.speaker.Rate = 0  # 음성 속도 (기본값)
            self.speaker.Volume = 100  # 음량 최대
            
            # 워커 스레드 시작
            self._start_worker_thread()
            
        except Exception as e:
            self.speaker = None
    
    def _start_worker_thread(self):
        """TTS 처리를 위한 워커 스레드를 시작합니다."""
        if self._worker_thread is None or not self._worker_thread.is_alive():
            self._is_running = True
            self._worker_thread = threading.Thread(target=self._tts_worker, daemon=True)
            self._worker_thread.start()
    
    def _tts_worker(self):
        """TTS 작업을 처리하는 워커 스레드"""
        while self._is_running:
            try:
                (name, rate), current_time = self._tts_queue.get(timeout=0.1)
                # 중복 제거 - 큐에서 가져온 후 다시 확인
                if (self.last_spoken_name == name and 
                    current_time - self.last_speak_time < 0.3):
                    self._tts_queue.task_done()
                    continue
                
                # TTS 처리
                if self.speaker:
                    to_speak = self._process_name_for_speech(name)
                    if to_speak:
                        try:
                            original_rate = self.speaker.Rate
                            if rate is not None:
                                self.speaker.Rate = rate
                            self.speaker.Speak(to_speak)
                            self.speaker.Rate = original_rate
                            self.last_speak_time = current_time
                            self.last_spoken_name = name
                        except Exception:
                            pass
                
                self._tts_queue.task_done()
                
            except Empty:
                # 타임아웃 - 계속 대기
                continue
            except Exception:
                # 기타 오류 - 무시하고 계속
                continue
    
    def _process_name_for_speech(self, name):
        """이름을 TTS용으로 처리합니다 (캐싱 사용)"""
        if not name:
            return None
        
        # 캐시에서 확인
        if name in self._name_cache:
            return self._name_cache[name]
        
        # 숫자(정수/실수)만으로 이루어진 경우 전체 읽기
        try:
            float(str(name))
            is_number = True
        except ValueError:
            is_number = False
        if is_number:
            result = name
        # 영어가 포함되어 있는지 확인 (정규식 사용)
        elif self._english_pattern.search(name):
            result = "영어"
        else:
            # 마지막 글자만 발음
            result = name[-1] if name else ""
        
        # 캐시 크기 제한 (메모리 사용량 조절)
        if len(self._name_cache) > 100:
            # 가장 오래된 항목 일부 제거
            items_to_remove = list(self._name_cache.keys())[:20]
            for key in items_to_remove:
                del self._name_cache[key]
        
        self._name_cache[name] = result
        return result
    
    def speak_name(self, name, rate=None):
        """이름(또는 숫자)을 음성으로 읽습니다 (비동기 처리, 속도 조절 가능)"""
        if not self.speaker or not name:
            return
            
        current_time = time.time()
        
        # 같은 이름이면 0.3초 이내 중복 호출 방지
        if (self.last_spoken_name == name and 
            current_time - self.last_speak_time < 0.3):
            return
        
        # 워커 스레드가 실행 중이 아니면 시작
        if not self._is_running or not self._worker_thread.is_alive():
            self._start_worker_thread()
        
        # 큐가 가득 찬 경우 가장 오래된 항목 제거
        try:
            while self._tts_queue.full():
                try:
                    self._tts_queue.get_nowait()
                    self._tts_queue.task_done()
                except Empty:
                    break
            
            # 새 TTS 작업 추가 (rate도 함께 전달)
            self._tts_queue.put_nowait(((name, rate), current_time))
            
        except Exception:
            # 큐 작업 실패 시 무시
            pass
    
    def stop(self):
        """TTS 매니저를 정지합니다."""
        self._is_running = False
        if self._worker_thread and self._worker_thread.is_alive():
            self._worker_thread.join(timeout=1.0)
        
        # 큐 정리
        try:
            while not self._tts_queue.empty():
                self._tts_queue.get_nowait()
                self._tts_queue.task_done()
        except Empty:
            pass
    
    def __del__(self):
        """소멸자에서 정리 작업 수행"""
        try:
            self.stop()
        except Exception:
            pass