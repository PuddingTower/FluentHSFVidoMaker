import cv2
import os
import sys

def create_video_from_images(frame_duration, output_filename='output.mp4'):
    # 获取当前程序的目录位置
    if getattr(sys, 'frozen', False):
        current_dir = os.path.dirname(sys.executable)
    else:
        current_dir = os.path.dirname(os.path.abspath(__file__))

    # 获取当前目录中的所有图片文件
    image_files = [f for f in os.listdir(current_dir) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
    image_files.sort()  # 根据文件名排序，确保顺序正确

    # 检查是否找到图片
    if len(image_files) == 0:
        print("No images found in the directory.")
        return

    # 获取第一张图片的大小，用于设置视频的尺寸
    first_image_path = os.path.join(current_dir, image_files[0])
    first_image = cv2.imread(first_image_path)
    if first_image is None:
        print(f"Failed to load the first image: {image_files[0]}")
        return

    height, width, layers = first_image.shape

    # 设置视频输出参数
    video_path = os.path.join(current_dir, output_filename)
    video = cv2.VideoWriter(video_path, cv2.VideoWriter_fourcc(*'mp4v'), 1 / frame_duration, (width, height))

    # 逐一添加每一帧到视频中
    for image_file in image_files:
        image_path = os.path.join(current_dir, image_file)
        img = cv2.imread(image_path)
        if img is None:
            print(f"Failed to load image: {image_file}")
            continue
        
        # 如果图像尺寸不匹配，则调整大小
        if img.shape[0] != height or img.shape[1] != width:
            print(f"Resizing image: {image_file}")
            img = cv2.resize(img, (width, height))

        video.write(img)

    video.release()
    print(f"Video processing completed. The video is saved at {video_path}")

if __name__ == "__main__":
    try:
        # 询问用户输入每一帧的持续时间
        frame_duration = float(input("Please enter the duration for each image frame (in seconds, e.g., 0.1): "))
        create_video_from_images(frame_duration, output_filename='output.mp4')
    except Exception as e:
        print(f"An error occurred: {e}")
    
    input("Press Enter to exit...")
