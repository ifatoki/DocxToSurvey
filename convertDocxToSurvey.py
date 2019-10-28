import docx
import re
import argparse
from datetime import datetime

parser = argparse.ArgumentParser()
parser.add_argument('--source-file-path', action='store', required=True)
parser.add_argument('--output-file-path', action='store', required=False)
args = parser.parse_args()

def convertCurrentUnit(currentUnit):
  def newPageTransform(text):
    return '{} {}'.format('::NewPage::', text[:-9].strip())

  def radioButtonTransform(options):
    returnArray = []
    for option in options:
      returnArray = [*returnArray, '() {}'.format(option)]
    return '\n'.join(returnArray)

  def textBoxTransform(text):
    return '{}\n_'.format(text[:-10].strip())

  def essayTransform(text):
    return '{}\n_\n_'.format(text[:-7].strip())

  def checkBoxTransform(options):
    returnArray = []
    for option in options:
      returnArray = [*returnArray, '[] {}'.format(option)]
    return '\n'.join(returnArray)

  def tableTransform(placeholder, options):
    returnArray = []
    i = 0
    if re.findall('[.+]', options[-1]):
      headers = options.pop()[4:-1]
      headers = re.split('[,;]?\s\d\.\s', headers)
    else:
      headers = ['values']
    postscript = '\t'.join([ placeholder for header in headers ])
    returnArray = [ ' {}'.format('\t'.join(headers)) ]
    while i < len(options):
      returnArray = [*returnArray, '{}\t{}'.format(options[i], postscript)]
      i += 1
    return '\n'.join(returnArray)

  output = []
  i = 0
  while i < len(currentUnit):
    line = re.sub('\[.*\]', '', currentUnit[i])
    if line.lower().endswith('(newpage)'):
      if len(output)>0: output[0] = '{}\n'.format(output[0])
      output = [*output, newPageTransform(line)]
      break
    elif line.lower().endswith('(rb)'):
      output = ['{} {}'.format(' '.join(output), line[:-14].strip()).strip(), radioButtonTransform(currentUnit[i+1:])]
      break
    elif line.lower().endswith('(tb)'):
      output = [*output, textBoxTransform(line)]
      break
    elif line.lower().endswith('(e)'):
      output = [*output, essayTransform(line)]
      break
    elif line.lower().endswith('(cb)'):
      output = ['{} {}'.format(' '.join(output), line[:-11].strip()).strip(), checkBoxTransform(currentUnit[i+1:])]
      break
    elif line.lower().endswith('(rt)'):
      output = ['{} {}'.format(' '.join(output), line[:-13].strip()).strip(), tableTransform('()', currentUnit[i+1:])]
      break
    elif line.lower().endswith('(ct)'):
      output = ['{} {}'.format(' '.join(output), line[:-16].strip()).strip(), tableTransform('[]', currentUnit[i+1:])]
      break
    elif line.lower().endswith('(tt)'):
      output = ['{} {}'.format(' '.join(output), line[:-12].strip()).strip(), tableTransform('_', currentUnit[i+1:])]
      break
    else:
      output = [*output, line]
    i += 1

  return output

def main():
  doc = docx.Document(args.source_file_path)
  currentUnit = []
  outputFilePath = args.output_file_path
  if not outputFilePath: outputFilePath = '{}_output_{}.txt'.format(args.source_file_path[:-5],datetime.now().strftime('%d%m%Y%H%M'))
  outputFile = open(outputFilePath, 'w')
  for paragraph in doc.paragraphs:
    if paragraph.text.strip() == '':
      if (len(currentUnit)>0):
        outputFile.write('\n'.join(convertCurrentUnit(currentUnit)))
        outputFile.write('\n\n')
      currentUnit = []
    else:
      currentUnit = [*currentUnit, paragraph.text.strip()]
  outputFile.close()

main()