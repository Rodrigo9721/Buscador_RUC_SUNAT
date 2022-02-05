# Buscador_RUC_SUNAT

https://pylessons.com/TensorFlow-CAPTCHA-solver-training/

Entrenamiento de la red neuronal:

Descargar las imágenes y etiquetarlas (https://github.com/tzutalin/labelImg)
Añadir las imágenes con sus archivos xlm en la carpeta CAPTCHA_images\train y CAPTCHA_images\test de forma aleatoria.
Correr el archivo xml_to_csv.py
Editar los paths en las líneas 15-17 en generate_tfrecord.py para train y test, y correr el código (una vez para train y otra para test).
Para empezar el entrenamiento poner: python train.py --logtostderr --train_dir=CAPTCHA_training_dir/ --pipeline_config_path=CAPTCHA_training/faster_rcnn_inception_v2_coco.config
en la terminal. Para visualizar el entrenamiento en tensorboard poner: tensorboard --logdir="C:\Users\Usuario\Anaconda3\envs\T\Lib\site-packages\tensorflow\models\research\object_detection\CAPTCHA_training_dir"
en otra terminal
Cuando la pérdida en la predicción ya no disminuya con cada step, poner en la terminal: python export_inference_graph.py --input_type image_tensor --pipeline_config_path CAPTCHA_training/faster_rcnn_inception_v2_coco.config --trained_checkpoint_prefix CAPTCHA_training_dir/model.ckpt-XXXXX --output_directory CAPTCHA_inference_graph
(reemplazando las XXXXX por el número que aparece en el archivo model.ckpt en la carpeta CAPTCHA_training_dir)
Para realizar las predicciones copiar el frozen_inference_graph como CAPTCHA_frozen_inference_graph y el labelmap.pbtxt como CAPTCHA_labelmap.pbtxt en el mismo directorio que el archivo CAPTCHA_object_detection.py y crear un nuevo archivo .py donde se importan las funciones de CAPTCHA_object_detection.py (se utiliza la función Captcha_detection('imagen.jpg'))

Si se quiere entrenar el modelo utilizando otras etiquetas, editar los archivos faster_rcnn_inception_v2_coco.config, labelmap.pbtxt en CAPTCHA_training y en generate_tfrecord.py
