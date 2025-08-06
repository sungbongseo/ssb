import subprocess
import webbrowser
import time
import os
import sys
import socket

def get_local_ip():
    """로컬 IP 주소 가져오기"""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except:
        return "localhost"

def check_python():
    """Python 설치 확인"""
    try:
        version = sys.version_info
        print(f"✓ Python {version.major}.{version.minor}.{version.micro} 감지됨")
        if version.major < 3 or (version.major == 3 and version.minor < 7):
            print("⚠ Python 3.7 이상 버전이 필요합니다!")
            return False
        return True
    except:
        print("❌ Python이 설치되어 있지 않습니다!")
        return False

def install_packages():
    """필요한 패키지 설치"""
    print("\n📦 필요한 패키지를 확인합니다...")
    
    packages = {
        'streamlit': 'streamlit',
        'pandas': 'pandas',
        'plotly': 'plotly',
        'openpyxl': 'openpyxl',
        'numpy': 'numpy'
    }
    
    for package_name, import_name in packages.items():
        try:
            __import__(import_name)
            print(f"✓ {package_name} 설치됨")
        except ImportError:
            print(f"📥 {package_name} 설치 중...")
            try:
                subprocess.check_call([
                    sys.executable, '-m', 'pip', 'install', 
                    package_name, '--quiet'
                ])
                print(f"✓ {package_name} 설치 완료")
            except:
                print(f"❌ {package_name} 설치 실패")
                return False
    
    return True

def check_dashboard_file():
    """대시보드 파일 확인"""
    if not os.path.exists('hr_dashboard.py'):
        print("❌ hr_dashboard.py 파일을 찾을 수 없습니다!")
        print("   현재 폴더:", os.getcwd())
        return False
    print("✓ 대시보드 파일 확인됨")
    return True

def kill_existing_streamlit():
    """기존 Streamlit 프로세스 종료"""
    try:
        # Windows
        if sys.platform == "win32":
            subprocess.call(['taskkill', '/F', '/IM', 'streamlit.exe'], 
                          stdout=subprocess.DEVNULL, 
                          stderr=subprocess.DEVNULL)
    except:
        pass

def run_dashboard():
    """메인 실행 함수"""
    print("=" * 50)
    print("🎯 HR 대시보드 실행 프로그램")
    print("=" * 50)
    
    # 1. Python 확인
    if not check_python():
        input("\n엔터를 눌러 종료하세요...")
        return
    
    # 2. 대시보드 파일 확인
    if not check_dashboard_file():
        input("\n엔터를 눌러 종료하세요...")
        return
    
    # 3. 패키지 설치
    if not install_packages():
        print("\n⚠ 일부 패키지 설치에 실패했습니다.")
        print("수동으로 설치해보세요: pip install -r requirements.txt")
        input("\n엔터를 눌러 종료하세요...")
        return
    
    # 4. 기존 프로세스 종료
    kill_existing_streamlit()
    
    # 5. Streamlit 실행
    print("\n🚀 대시보드를 시작합니다...")
    local_ip = get_local_ip()
    
    print(f"\n📌 접속 주소:")
    print(f"   - 본인 PC: http://localhost:8501")
    print(f"   - 다른 PC: http://{local_ip}:8501")
    print("\n⚠ 종료하려면 이 창을 닫거나 Ctrl+C를 누르세요")
    print("=" * 50)
    
    # Streamlit 실행
    process = subprocess.Popen([
        sys.executable, '-m', 'streamlit', 'run',
        'hr_dashboard.py',
        '--server.address', '0.0.0.0',
        '--server.port', '8501',
        '--server.headless', 'true',
        '--browser.gatherUsageStats', 'false'
    ])
    
    # 3초 후 브라우저 열기
    time.sleep(3)
    webbrowser.open('http://localhost:8501')
    
    try:
        # 프로세스 대기
        process.wait()
    except KeyboardInterrupt:
        print("\n\n👋 대시보드를 종료합니다...")
        process.terminate()

if __name__ == "__main__":
    try:
        run_dashboard()
    except Exception as e:
        print(f"\n❌ 오류 발생: {e}")
        input("\n엔터를 눌러 종료하세요...")