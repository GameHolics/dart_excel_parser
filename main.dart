import 'dart:io';

import 'package:excel_parser/excel_parser.dart' as excel_parser;
import 'package:args/args.dart' as args;

const availableFormats = ['json', 'csv', 'tsv'];

Future<void> main(List<String> arguments) async {
  final argsParser = args.ArgParser();
  argsParser.addFlag('help', abbr: 'h', negatable: false, help: 'Show usage');
  argsParser.addOption(
    'format',
    abbr: 'f',
    defaultsTo: 'json',
    allowed: availableFormats,
  );
  argsParser.addFlag('prettify', abbr: 'p', negatable: false, help: 'Prettify json output');
  argsParser.addOption('outFile', abbr: 'o', help: 'If not provided, will print to console.');

  try {
    final results = argsParser.parse(arguments);

    if (results['help'] || results.arguments.isEmpty) {
      print(argsParser.usage);
      exit(0);
    }

    final rests = results.rest;
    if (rests.isEmpty) {
      print('No file to parsed!');
      exit(1);
    }

    final targetFileName = rests[0];

    final targetFile = File(targetFileName);

    if (!await targetFile.exists() || !excel_parser.canParse(targetFileName)) {
      print('$targetFileName is not a valid file!');
      exit(1);
    }

    bool usePrettify = results['prettify'];

    File? outputFile;
    if (results['outFile'] != null) {
      final outputFileName = (results['outFile'] as String);
      if (outputFileName.isNotEmpty) {
        outputFile = File(outputFileName);
      }
    }

    final bytes = await targetFile.readAsBytes();

    final rawData = excel_parser.parseRaw(bytes);

    var output = '';
    switch (results['format'] as String) {
      case 'json':
        output = excel_parser.rawDataToJson(rawData, usePrettify);
        break;
      case 'csv':
        output = excel_parser.rawDataToSeparatedText(rawData, ',');
        break;
      case 'tsv':
        output = excel_parser.rawDataToSeparatedText(rawData, '\t');
        break;
    }

    if (outputFile != null) {
      outputFile = await outputFile.create(recursive: true);
      await outputFile.writeAsString(output, mode: FileMode.write, flush: true);
    } else {
      print(output);
    }
  } catch (e) {
    print(argsParser.usage);
    print(e);
    rethrow;
  }
}
