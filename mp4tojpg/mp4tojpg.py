import cv2

def extract_frame(video_path, frame_number, output_path):
    # ビデオファイルを読み込む
    cap = cv2.VideoCapture(video_path)
    
    # フレーム番号を設定
    cap.set(cv2.CAP_PROP_POS_FRAMES, frame_number)
    
    # フレームを読み込む
    success, frame = cap.read()
    
    if success:
        # フレームの読み込みが成功した場合、画像として保存
        cv2.imwrite(output_path, frame)
        print(f"Frame {frame_number} has been saved to {output_path}")
    else:
        # フレームの読み込みに失敗した場合
        print(f"Failed to extract frame at {frame_number}")
    
    # リソースを解放
    cap.release()

# 使用例
video_path = 'path/to/your/video.MOV'
frame_number = 100  # 抽出したいフレーム番号
output_path = 'output_frame.jpg'  # 保存するファイル名

extract_frame(video_path, frame_number, output_path)