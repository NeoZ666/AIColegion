import cv2
import openpyxl

name_list = []

def cleanup_data():
    with open("names.txt") as file:
        for line in file:
            name_list.append(line.strip())

def generate_certi():
    for name in name_list:
        template = cv2.imread("certi_template.jpg")
        if len(name) > 10:
            cv2.putText(template, name, (385, 515), cv2.FONT_HERSHEY_TRIPLEX, 2, (255,255,0), 2, cv2.LINE_AA)
        else:
            cv2.putText(template, name, (445, 515), cv2.FONT_HERSHEY_TRIPLEX, 2, (255, 255, 0), 2, cv2.LINE_AA)
        cv2.imwrite(f'generated_certi_data/{name}.jpg',template)


cleanup_data()
generate_certi()