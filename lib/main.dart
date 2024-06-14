import 'dart:io';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:flutter_barcode_scanner/flutter_barcode_scanner.dart';
import 'package:excel/excel.dart';
import 'package:path/path.dart' as p;
import 'package:share_plus/share_plus.dart';
import 'package:cross_file/cross_file.dart';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      theme: ThemeData.dark(),
      home: const MainApp(),
    );
  }
}

class MainApp extends StatefulWidget {
  const MainApp({super.key});

  @override
  State<MainApp> createState() => _MyAppState();
}

class _MyAppState extends State<MainApp> {
  String? scanResult;
  List<String> rollNumbers = [];
  TextEditingController _fileNameController = TextEditingController();

  Future<String?> scanBar() async {
    try {
      scanResult = await FlutterBarcodeScanner.scanBarcode(
        "#ff6666",
        "Cancel",
        true,
        ScanMode.BARCODE,
      );
    } on PlatformException {
      scanResult = "Failed to get platform version";
    }
    setState(() {
      rollNumbers.add(scanResult ?? "");
    });
    return scanResult;
  }

  void rollNumberRemove(int index) {
    setState(() {
      rollNumbers.removeAt(index);
    });
  }

  void rollNumberClear() {
    setState(() {
      rollNumbers.clear();
    });
  }

  Future<void> _showFileNameDialog(BuildContext context) async {
    return showDialog<void>(
      context: context,
      barrierDismissible: false, // User must tap a button!
      builder: (BuildContext context) {
        return AlertDialog(
          title: const Text('Enter name for file'),
          content: TextField(
            controller: _fileNameController,
            decoration: const InputDecoration(hintText: "Enter name"),
          ),
          actions: <Widget>[
            TextButton(
              child: const Text('Cancel'),
              onPressed: () {
                Navigator.of(context).pop();
              },
            ),
            TextButton(
              child: const Text('OK'),
              onPressed: () {
                Navigator.of(context).pop();
                exportFile(_fileNameController.text);
              },
            ),
          ],
        );
      },
    );
  }

  Future<void> exportFile(String fileName) async {
    // 1. Create an Excel instance
    final excel = Excel.createExcel();

    // 2. Create a new sheet
    final sheet = excel['Sheet1'];

    // 3. Write data from the list to the sheet, starting from A1
    int rowIndex = 1;
    for (final rollNumber in rollNumbers) {
      sheet.cell(CellIndex.indexByString('A$rowIndex')).value =
          TextCellValue(rollNumber);
      rowIndex++;
    }
    var fileBytes = excel.save();

    String filePath = p.join('/storage/emulated/0/Download', '$fileName.xlsx');

    File(filePath)
      ..createSync(recursive: true)
      ..writeAsBytesSync(fileBytes!);

    final result = await Share.shareXFiles([XFile(filePath)],
        text: 'Excel created please share');

    print('Excel sheet shared successfully.$result');
  }

  Widget buildListTile(String rollNumber) {
    return ListTile(
      contentPadding: const EdgeInsets.all(10.0),
      shape: RoundedRectangleBorder(
        side: const BorderSide(
            color: Color.fromARGB(255, 86, 86, 86), width: 0.5),
        borderRadius: BorderRadius.circular(50),
      ),
      leading: const Icon(Icons.check_circle_outline),
      title: Text(
        rollNumber,
        style: const TextStyle(fontWeight: FontWeight.bold),
      ),
      trailing: IconButton(
        icon: const Icon(Icons.delete),
        onPressed: () => rollNumberRemove(rollNumbers.indexOf(rollNumber)),
      ),
    );
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text("Barcode/QR Code to Excel"),
        centerTitle: true,
      ),
      body: Center(
        child: Column(
          mainAxisAlignment: MainAxisAlignment.center,
          children: [
            Expanded(
              child: ListView.builder(
                itemCount: rollNumbers.length,
                itemBuilder: (context, index) {
                  final rollNumber = rollNumbers[index];
                  return buildListTile(rollNumber);
                },
              ),
            ),
            const SizedBox(height: 20),
            Align(
              alignment: Alignment.bottomCenter,
              child: Container(
                height: 60.0,
                width: double.infinity,
                color: Colors.transparent,
                child: Row(
                  mainAxisAlignment: MainAxisAlignment.center,
                  children: [
                    FloatingActionButton.extended(
                      onPressed: () async {
                        await scanBar();
                      },
                      icon: const Icon(Icons.camera_alt_outlined),
                      label: const Text("Scan"),
                    ),
                    const SizedBox(width: 20.0),
                    FloatingActionButton.extended(
                      onPressed: () async {
                        rollNumberClear();
                      },
                      icon: const Icon(Icons.delete_outline),
                      label: const Text("Clear"),
                    ),
                    const SizedBox(width: 20.0),
                    FloatingActionButton.extended(
                      onPressed: () async {
                        await _showFileNameDialog(context);
                      },
                      icon: const Icon(Icons.ios_share_outlined),
                      label: const Text("Export"),
                    ),
                  ],
                ),
              ),
            ),
            const SizedBox(
              height: 10,
            ),
            const Text("Developed by Geoseiden",
                style: TextStyle(
                  fontSize: 12, // Adjust font size as needed
                  color: Colors.grey,
                ))
          ],
        ),
      ),
    );
  }
}
