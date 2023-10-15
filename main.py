import argparse
import logging
from doc_translate.translate import init_model, get_layout, translate_layout
from doc_translate.utils.file_utils import convert_info_docx

def main(args):
    model = init_model(args.model_name)
    layout = get_layout(args.input_file, model, args.file_type)
    translated_layout = translate_layout(layout, args.source_lang, args.target_lang)
    convert_info_docx(translated_layout, args.output_dir, args.output_file)
    logging.info("Translated")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Translate layout file')
    parser.add_argument('--input_file', type=str, help='Input file path')
    parser.add_argument('--output_dir', type=str, help='Output directory path')
    parser.add_argument('--output_file', type=str, help='Output file name')
    parser.add_argument('--source_lang', type=str, help='Source language')
    parser.add_argument('--target_lang', type=str, help='Target language')
    parser.add_argument('--model_name', type=str, default='chipperv1', help='Model name')
    parser.add_argument('--file_type', type=str, default='pdf', help='File type')
    args = parser.parse_args()
    main(args)