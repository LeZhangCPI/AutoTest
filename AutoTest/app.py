import os
import subprocess
import sys
from flask import Flask, request, jsonify
from flask_cors import CORS
import os
app1 = Flask(__name__, static_folder='dist')
CORS(app1)

@app1.route('/AutoTest', methods=['POST'])      #notation
def run_patent_script():
    data = request.json
    entity_status = data.get('entityStatus')

    if entity_status == 'Reports':
        try:
            subprocess.run([sys.executable, '/Users/lez/Desktop/AutoTest/AutoTest/DueDateList.py'],
                           check=True, env=os.environ.copy())
            return jsonify({'message': 'DueDateListWithExcel.py script executed successfully'}), 200
        except subprocess.CalledProcessError as e:
            return jsonify({'error': str(e)}), 500
    elif entity_status == 'Country':
        try:
            subprocess.run([sys.executable, '/Users/lez/Desktop/AutoTest/AutoTest/Country.py'],
                           check=True, env=os.environ.copy())
            return jsonify({'message': 'Country.py script executed successfully'}), 200
        except subprocess.CalledProcessError as e:
            return jsonify({'error': str(e)}), 500
    else:
        return jsonify({'message': 'No script executed'}), 200


if __name__ == '__main__':
    app1.run(debug=True)