(defun c:getAllTextWithCoords (/ content obj-name) 
  (setq doc (vla-get-ActiveDocument (vlax-get-acad-object))) ; 获取当前文档
  (setq ms (vla-get-ModelSpace doc)) ; 获取模型空间

  (princ "\n图框内所有文本及其坐标：") ; 打印标题
  (setq content "数量文本, 坐标\n")
  (vlax-for entity ms  ; 遍历模型空间中的所有实体
    (progn 
      (setq obj-name (vla-get-objectname entity))
      (if (or (equal "AcDbText" obj-name)  ; 检查是否为文本对象
        (equal "AcDbMText" obj-name)) ; 检查是否为多行文本对象
        (progn
          (setq text-item (vla-get-textstring entity)) ; 获取文本字符串
          (if (and text-item (filterStrWithX text-item))
            (progn
              (setq text-item (vla-get-textstring entity)) ; 获取文本字符串
              (setq position (vlax-safeArray->list (vlax-variant-value (vla-get-insertionpoint entity)))) ; 获取实体插入点坐标
              (setq x (rtos (car position) 2 2))
              (setq y (rtos (cadr position) 2 2))
              (setq z (rtos (caddr position) 2 2))
              (setq row (strcat text-item ",\"(" x ", " y ", " z ")\"\n"))
              (princ (strcat "信息：" obj-name "," row ))
              (setq content (strcat content row))
            )
          ) 
        )        
      )
    )
  )
  (princ content)
  ; 保存文件
  (savefile "e:\\temp\\numbers.csv" content)
  (princ) ; 结束
)

(defun filterStrWithX (str) 
  ;; 判断字符串是否以"X"开头
  (if 
    (and (not (null str))  ; 确保字符串不为空
         (or (= (substr str 1 1) "x") 
             (= (substr str 1 1) "X")
             (= (substr str 1 1) "*")
         )
    )
    T ; 如果以"X"开头，返回真
    nil ; 否则返回假
  )
)


(defun c:getAllDeviceBlocks (/ ss num ss-name insert-name block-ref content devName devType) 
  (setq ss (ssget "X" 
                  '((0 . "INSERT")
                    (-4 . "<or")
                    (2 . "设备 RS485(XT)")
                    (2 . "33M")
                    (-4 . "or>")
                   )
           )
  )
  (setq num 0) ; 使用num 循环遍历 图框选择集
  (setq content "设备类型, 回路号, 坐标\n")
  (repeat (sslength ss) 
    (setq ss-name (ssname ss num))
    (setq block-ref (vlax-ename->vla-object ss-name))
    (setq attrs (vlax-safeArray->list (vlax-variant-value (vla-getattributes block-ref))))
    (setq devType "-")
    (setq devName "-")
    (foreach attr attrs
      (setq tag (vlax-get-property attr 'TagString))
      (setq name (vlax-get-property attr 'TextString))
      (cond
        ((= tag "设备名称") (setq devType name))
        ((= tag "开关柜名称") (setq devName name))
      )
      )
    (setq position (vlax-safeArray->list (vlax-variant-value (vla-get-insertionpoint block-ref))))
    (setq x (rtos (car position) 2 2))
    (setq y (rtos (cadr position) 2 2))
    (setq z (rtos (caddr position) 2 2))
    (setq content (strcat content (strcat devType "," devName ",\"(" x ", " y ", " z ")\"\n")))
    (setq num (1+ num))
  )
  (princ content)
  ; 保存文件
  (savefile "e:\\temp\\devices.csv" content)
  (princ)
)

(defun saveFile (filename content / file) 
  (setq file (open filename "w")) ; 打开文件，写模式
  (if file 
    (progn 
      (write-line content file) ; 写入文本
      (close file) ; 关闭文件
      (princ (strcat "内容已写入 " filename))
    )
    (princ "文件打开失败！")
  )
  (princ)
)







