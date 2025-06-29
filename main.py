from PySide6.QtWidgets import QApplication
from PySide6.QtCore import QCoreApplication
from ui.main_window import MainWindow
from core.score_logic import ScoreLogic
from services.tts_manager import TTSManager
import sys
import traceback
import gc
import atexit

def cleanup_resources():
    """애플리케이션 종료 시 리소스 정리"""
    try:
        # TTS 매니저 정리
        if hasattr(cleanup_resources, 'tts_manager'):
            cleanup_resources.tts_manager.stop()
        
        # 가비지 컬렉션 강제 실행
        gc.collect()
    except Exception:
        pass

def setup_application():
    """애플리케이션 초기 설정 최적화"""
    app = QApplication(sys.argv)
    
    # 애플리케이션 속성 최적화 (PySide6에서 지원하지 않는 속성은 주석 처리)
    # app.setAttribute(app.AA_DontShowIconsInMenus, False)  # 지원되지 않음
    # app.setAttribute(app.AA_NativeWindows, False)         # 지원되지 않음
    
    # 스타일 및 성능 최적화 (PySide6에서 지원하지 않는 속성은 주석 처리)
    # app.setEffectEnabled(app.UI_AnimateMenu, False)
    # app.setEffectEnabled(app.UI_AnimateCombo, False)
    # app.setEffectEnabled(app.UI_AnimateTooltip, False)
    
    # 애플리케이션 정보 설정
    QCoreApplication.setApplicationName("수행평가 점수 입력기")
    QCoreApplication.setApplicationVersion("1.0")
    QCoreApplication.setOrganizationName("melderse 짐승농장")
    
    return app

def create_components():
    """컴포넌트 생성 및 초기화"""
    try:
        # 로직 컴포넌트 생성
        logic = ScoreLogic()
        
        # TTS 매니저 생성 (싱글톤)
        tts = TTSManager()
        
        # 정리 함수에서 참조할 수 있도록 저장
        cleanup_resources.tts_manager = tts
        
        # 메인 윈도우 생성
        window = MainWindow(logic, tts)
        
        return logic, tts, window
    
    except Exception as e:
        print(f"컴포넌트 생성 중 오류 발생: {e}")
        print(traceback.format_exc())
        return None, None, None

def main():
    """메인 함수 - 최적화된 애플리케이션 실행"""
    app = None
    window = None
    
    try:
        # 정리 함수 등록
        atexit.register(cleanup_resources)
        
        # 애플리케이션 설정
        app = setup_application()
        
        # 컴포넌트 생성
        logic, tts, window = create_components()
        
        if window is None:
            print("윈도우 생성 실패")
            return 1
        
        # 윈도우 표시
        window.show()
        
        # 초기 가비지 컬렉션
        gc.collect()
        
        # 이벤트 루프 실행
        exit_code = app.exec()
        
        return exit_code
        
    except KeyboardInterrupt:
        print("사용자에 의해 프로그램이 중단되었습니다.")
        return 0
        
    except Exception as e:
        print("예상치 못한 오류 발생:")
        print(str(e))
        print("\n상세 추적 정보:")
        print(traceback.format_exc())
        return 1
        
    finally:
        # 명시적 리소스 정리
        try:
            if window:
                window.close()
                window.deleteLater()
            
            if app:
                app.quit()
                app.deleteLater()
            
            # 최종 가비지 컬렉션
            gc.collect()
            
        except Exception:
            # 정리 중 오류는 무시
            pass

if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)