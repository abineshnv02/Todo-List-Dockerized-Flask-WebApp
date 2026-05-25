from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Create a new Document
doc = Document()

# Add title
title = doc.add_heading('Todo List Dockerized Flask WebApp', 0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER

subtitle = doc.add_heading('Complete Study Guide for Python Developer Role Interview', level=2)
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_paragraph()

# PART 1
doc.add_heading('PART 1: PROJECT OVERVIEW', level=1)
doc.add_paragraph(
    'This is a simple but complete real-world project that demonstrates:',
    style='List Bullet'
)
doc.add_paragraph('Backend: Python Flask web framework', style='List Bullet 2')
doc.add_paragraph('Database: SQLite with SQLAlchemy ORM', style='List Bullet 2')
doc.add_paragraph('Frontend: HTML/Jinja2 templating with Semantic UI CSS', style='List Bullet 2')
doc.add_paragraph('Containerization: Docker for deployment', style='List Bullet 2')
doc.add_paragraph('DevOps: Docker Compose for orchestration and Volume mapping for data persistence', style='List Bullet 2')

doc.add_paragraph()
doc.add_paragraph('Key Skills Demonstrated:', style='Heading 3')
doc.add_paragraph('Full-stack web development', style='List Bullet')
doc.add_paragraph('Relational databases and ORM', style='List Bullet')
doc.add_paragraph('Containerization and deployment', style='List Bullet')
doc.add_paragraph('Data persistence strategies', style='List Bullet')

# PART 2
doc.add_heading('PART 2: ARCHITECTURE & TECHNOLOGY STACK', level=1)

doc.add_heading('2.1 Backend Technologies', level=2)
doc.add_paragraph('Flask (3.0.0) → Web framework', style='List Bullet')
doc.add_paragraph('Flask-SQLAlchemy (3.1.1) → Database ORM (Object-Relational Mapping)', style='List Bullet')
doc.add_paragraph('SQLAlchemy (2.2.7) → Database toolkit', style='List Bullet')
doc.add_paragraph('Gunicorn (21.2.0) → WSGI HTTP Server (for production)', style='List Bullet')
doc.add_paragraph('Werkzeug (3.0.1) → WSGI utility library', style='List Bullet')
doc.add_paragraph('Click (8.1.7) → CLI creation toolkit', style='List Bullet')

doc.add_heading('2.2 Application Structure', level=2)
doc.add_paragraph('app.py → Main Flask application (ALL backend logic)', style='List Bullet')
doc.add_paragraph('requirements.txt → Python dependencies', style='List Bullet')
doc.add_paragraph('Dockerfile → Container configuration', style='List Bullet')
doc.add_paragraph('docker-compose.yml → Multi-container orchestration', style='List Bullet')
doc.add_paragraph('db/ → Database storage location', style='List Bullet')
doc.add_paragraph('templates/base.html → Frontend HTML template', style='List Bullet')

# PART 3
doc.add_heading('PART 3: UNDERSTANDING THE DATABASE MODEL', level=1)

doc.add_heading('3.1 The Todo Class (Database Model)', level=2)
doc.add_paragraph(
    'class Todo(db.Model):\n'
    '    id = db.Column(db.Integer, primary_key=True)\n'
    '    task = db.Column(db.String(1000), nullable=False)\n'
    '    complete = db.Column(db.Boolean)\n'
    '    user_id = db.Column(db.Integer)',
    style='No Spacing'
)

doc.add_paragraph()
doc.add_paragraph('Key Concepts:', style='Heading 3')
doc.add_paragraph('Uses SQLAlchemy ORM (Object-Relational Mapping) - write Python classes instead of SQL', style='List Bullet')
doc.add_paragraph('db.Model: Base class for all database models', style='List Bullet')
doc.add_paragraph('Primary Key: id uniquely identifies each record', style='List Bullet')
doc.add_paragraph('Nullable: nullable=False means field is required', style='List Bullet')
doc.add_paragraph('Database: SQLite (file-based, lightweight, perfect for learning)', style='List Bullet')

# PART 4
doc.add_heading('PART 4: CORE FLASK FUNCTIONALITY', level=1)

doc.add_heading('4.1 Application Initialization', level=2)
doc.add_paragraph(
    'app = Flask(__name__)\n'
    'app.config[\'SQLALCHEMY_DATABASE_URI\'] = f\'sqlite:///{db_path}\'\n'
    'app.config[\'SQLALCHEMY_TRACK_MODIFICATIONS\'] = False\n'
    'db = SQLAlchemy(app)',
    style='No Spacing'
)

doc.add_heading('4.2 The Four Main Routes (Endpoints)', level=2)
table = doc.add_table(rows=5, cols=4)
table.style = 'Light Grid Accent 1'
hdr_cells = table.rows[0].cells
hdr_cells[0].text = 'Route'
hdr_cells[1].text = 'Method'
hdr_cells[2].text = 'Purpose'
hdr_cells[3].text = 'Logic'

rows_data = [
    ('/', 'GET', 'Display all tasks', 'Query all todos, render HTML template'),
    ('/add', 'POST', 'Add new task', 'Get form data, validate, create Todo, save to DB'),
    ('/delete/<id>', 'GET', 'Delete task', 'Find task by ID, delete from DB'),
    ('/update/<id>', 'GET', 'Toggle completion', 'Find task by ID, flip complete boolean, save'),
]

for i, (route, method, purpose, logic) in enumerate(rows_data, 1):
    row_cells = table.rows[i].cells
    row_cells[0].text = route
    row_cells[1].text = method
    row_cells[2].text = purpose
    row_cells[3].text = logic

doc.add_heading('4.3 Route Deep Dive', level=2)

doc.add_paragraph('Route 1: Home Page (/):', style='Heading 3')
doc.add_paragraph(
    '@app.route(\'/\')\n'
    'def index():\n'
    '    todoList = Todo.query.all()\n'
    '    return render_template(\'base.html\', todo_list=todoList)',
    style='No Spacing'
)
doc.add_paragraph('Query Database: Todo.query.all() fetches all Todo records', style='List Bullet')
doc.add_paragraph('Render Template: Sends data to HTML template for display', style='List Bullet')

doc.add_paragraph('Route 2: Add Task (/add):', style='Heading 3')
doc.add_paragraph(
    '@app.route(\'/add\', methods=["POST"])\n'
    'def add():\n'
    '    title = request.form.get("title")\n'
    '    if title == "":\n'
    '        return redirect(url_for("index"))\n'
    '    newTask = Todo(task=title, complete=False)\n'
    '    db.session.add(newTask)\n'
    '    db.session.commit()\n'
    '    return redirect(url_for("index"))',
    style='No Spacing'
)
doc.add_paragraph('POST Method: Secure way to submit form data', style='List Bullet')
doc.add_paragraph('Request Handling: Extract data from form', style='List Bullet')
doc.add_paragraph('ORM Operations: Create object → add to session → commit', style='List Bullet')
doc.add_paragraph('Redirect: Standard pattern after form submission', style='List Bullet')

doc.add_paragraph('Route 3: Delete Task (/delete/<todo_id>):', style='Heading 3')
doc.add_paragraph(
    '@app.route(\'/delete/<int:todo_id>\')\n'
    'def delete(todo_id):\n'
    '    task = Todo.query.filter_by(id=todo_id).first()\n'
    '    db.session.delete(task)\n'
    '    db.session.commit()\n'
    '    return redirect(url_for("index"))',
    style='No Spacing'
)
doc.add_paragraph('URL Parameters: <int:todo_id> captures the ID from URL', style='List Bullet')
doc.add_paragraph('Query Filtering: filter_by() finds specific record', style='List Bullet')
doc.add_paragraph('first(): Returns single record (not a list)', style='List Bullet')

doc.add_paragraph('Route 4: Update/Toggle Task (/update/<todo_id>):', style='Heading 3')
doc.add_paragraph(
    '@app.route(\'/update/<int:todo_id>\')\n'
    'def update(todo_id):\n'
    '    task = Todo.query.filter_by(id=todo_id).first()\n'
    '    task.complete = not task.complete\n'
    '    db.session.commit()\n'
    '    return redirect(url_for("index"))',
    style='No Spacing'
)

# PART 5
doc.add_heading('PART 5: FRONTEND - HTML TEMPLATING (Jinja2)', level=1)

doc.add_heading('5.1 What is Jinja2?', level=2)
doc.add_paragraph('Flask\'s template engine', style='List Bullet')
doc.add_paragraph('Allows embedding Python logic in HTML', style='List Bullet')
doc.add_paragraph('Syntax: {{ variable }} for variables, {% for loop %} for logic', style='List Bullet')

doc.add_heading('5.2 Key Template Features in base.html', level=2)

doc.add_paragraph('Form for Adding Tasks:', style='Heading 3')
doc.add_paragraph(
    '<form class="ui form" action="/add" method="POST">\n'
    '    <input type="text" name="title" placeholder="Enter Todo...">\n'
    '    <button type="submit">Add</button>\n'
    '</form>',
    style='No Spacing'
)
doc.add_paragraph('action="/add": Submits to /add route', style='List Bullet')
doc.add_paragraph('method="POST": Secure form submission', style='List Bullet')
doc.add_paragraph('name="title": Form field name (matches request.form.get("title"))', style='List Bullet')

doc.add_paragraph('Loop Through Todos:', style='Heading 3')
doc.add_paragraph(
    '{% for todo in todo_list %}\n'
    '    <p>{{ todo.id }} | {{ todo.task }}</p>\n'
    '    {% if todo.complete == False %}\n'
    '        <span>Not Complete</span>\n'
    '    {% else %}\n'
    '        <span>Completed</span>\n'
    '    {% endif %}\n'
    '    <a href="/update/{{ todo.id }}">Mark Finished</a>\n'
    '    <a href="/delete/{{ todo.id }}">Delete</a>\n'
    '{% endfor %}',
    style='No Spacing'
)
doc.add_paragraph('Jinja2 Syntax: {% %} for logic, {{ }} for variables', style='List Bullet')
doc.add_paragraph('Dynamic URLs: href="/delete/{{ todo.id }}" creates URLs dynamically', style='List Bullet')
doc.add_paragraph('Conditional Rendering: Show different text based on completion status', style='List Bullet')

doc.add_heading('5.3 CSS Framework', level=2)
doc.add_paragraph('Uses Semantic UI (CDN): Pre-built responsive CSS components', style='List Bullet')
doc.add_paragraph('Classes like ui button, ui form, ui segment provide styling', style='List Bullet')
doc.add_paragraph('No custom CSS needed - professional look out of the box', style='List Bullet')

# PART 6
doc.add_heading('PART 6: DATABASE & ORM DEEP DIVE', level=1)

doc.add_heading('6.1 What is an ORM (Object-Relational Mapping)?', level=2)

doc.add_paragraph('Without ORM (Raw SQL):', style='Heading 3')
doc.add_paragraph(
    'INSERT INTO todo (task, complete) VALUES (\'Buy milk\', 0);\n'
    'SELECT * FROM todo WHERE id = 1;\n'
    'DELETE FROM todo WHERE id = 1;',
    style='No Spacing'
)

doc.add_paragraph('With SQLAlchemy ORM (Python Objects):', style='Heading 3')
doc.add_paragraph(
    'newTask = Todo(task="Buy milk", complete=False)\n'
    'db.session.add(newTask)\n'
    'db.session.commit()\n'
    'task = Todo.query.filter_by(id=1).first()\n'
    'db.session.delete(task)\n'
    'db.session.commit()',
    style='No Spacing'
)

doc.add_paragraph('Benefits:', style='Heading 3')
doc.add_paragraph('Write Python instead of SQL', style='List Bullet')
doc.add_paragraph('Type safety', style='List Bullet')
doc.add_paragraph('Database-agnostic (switch from SQLite to PostgreSQL with minimal changes)', style='List Bullet')
doc.add_paragraph('SQL Injection protection', style='List Bullet')

doc.add_heading('6.2 CRUD Operations', level=2)
table2 = doc.add_table(rows=5, cols=3)
table2.style = 'Light Grid Accent 1'
hdr_cells2 = table2.rows[0].cells
hdr_cells2[0].text = 'Operation'
hdr_cells2[1].text = 'Method'
hdr_cells2[2].text = 'Example'

crud_data = [
    ('Create', 'Add + Commit', 'db.session.add(Todo(...))'),
    ('Read', 'Query', 'Todo.query.all() or .filter_by().first()'),
    ('Update', 'Modify + Commit', 'task.complete = True; db.session.commit()'),
    ('Delete', 'Delete + Commit', 'db.session.delete(task); db.session.commit()'),
]

for i, (operation, method, example) in enumerate(crud_data, 1):
    row_cells = table2.rows[i].cells
    row_cells[0].text = operation
    row_cells[1].text = method
    row_cells[2].text = example

doc.add_heading('6.3 SQLite Basics', level=2)
doc.add_paragraph('File-based database: Data stored in db/db.sqlite', style='List Bullet')
doc.add_paragraph('Lightweight: Perfect for learning, small projects', style='List Bullet')
doc.add_paragraph('Limitations: Not ideal for high-concurrency scenarios', style='List Bullet')
doc.add_paragraph('Data Persistence: Even after app restarts, data survives', style='List Bullet')

# PART 7
doc.add_heading('PART 7: CONTAINERIZATION WITH DOCKER', level=1)

doc.add_heading('7.1 What is Docker?', level=2)
doc.add_paragraph('Packages your application + all dependencies in a container (like a virtual machine) so it runs identically on any system.')

doc.add_heading('7.2 The Dockerfile Explained', level=2)
doc.add_paragraph(
    'FROM python:3.8-slim-buster\n'
    '# Base image: Ubuntu Linux with Python 3.8 pre-installed\n'
    '\n'
    'LABEL Maintainer_Name="Ahmed Ayman"\n'
    '# Metadata about the image\n'
    '\n'
    'WORKDIR /app\n'
    '# Sets the working directory inside container\n'
    '\n'
    'ENV FLASK_APP app.py\n'
    'ENV FLASK_ENV development\n'
    '# Environment variables\n'
    '\n'
    'COPY ./requirements.txt /requirements.txt\n'
    '# Copy dependencies file to container\n'
    '\n'
    'RUN pip3 install -r requirements.txt\n'
    '# Install Python packages inside container\n'
    '\n'
    'COPY . .\n'
    '# Copy entire project to container\n'
    '\n'
    'CMD python3 app.py\n'
    '# Command to run when container starts',
    style='No Spacing'
)

doc.add_paragraph('Key Concepts:', style='Heading 3')
doc.add_paragraph('Image: Blueprint/template (like a class)', style='List Bullet')
doc.add_paragraph('Container: Running instance (like an object)', style='List Bullet')
doc.add_paragraph('Layers: Each instruction creates a layer (caching for efficiency)', style='List Bullet')

doc.add_heading('7.3 Build & Run Process', level=2)
doc.add_paragraph(
    '# Build image from Dockerfile\n'
    'docker build -t todolist-flask:latest .\n'
    '\n'
    '# Create volume for database persistence\n'
    'docker volume create todolist.db\n'
    '\n'
    '# Run container\n'
    'docker run -d -p 5001:5000 -v todolist.db:/app/db todolist-flask',
    style='No Spacing'
)

doc.add_paragraph('Flags Explained:', style='Heading 3')
doc.add_paragraph('-d: Run in background (detached)', style='List Bullet')
doc.add_paragraph('-p 5001:5000: Map port 5001 (host) → 5000 (container)', style='List Bullet')
doc.add_paragraph('-v: Volume mounting (persist database data)', style='List Bullet')

doc.add_heading('7.4 Docker Compose', level=2)
doc.add_paragraph(
    'version: \'3.4\'\n'
    'services:\n'
    '  app:\n'
    '    image: todolistdockerizedflaskwebapp\n'
    '    build:\n'
    '      context: .\n'
    '      dockerfile: ./Dockerfile\n'
    '    ports:\n'
    '      - 80:5000\n'
    '    volumes:\n'
    '      - todolist.db:/app/db\n'
    'volumes:\n'
    '  todolist.db:',
    style='No Spacing'
)

doc.add_paragraph('Why Docker Compose?', style='Heading 3')
doc.add_paragraph('Single command to manage containers: docker-compose up', style='List Bullet')
doc.add_paragraph('Defines all configurations in one file', style='List Bullet')
doc.add_paragraph('Easy to reproduce environments', style='List Bullet')

# PART 8
doc.add_heading('PART 8: DEPLOYMENT & PRODUCTION', level=1)

doc.add_heading('8.1 Running Locally (Development)', level=2)
doc.add_paragraph(
    'python3 -m venv env\n'
    'source env/bin/activate\n'
    'pip install -r requirements.txt\n'
    'python3 app.py',
    style='No Spacing'
)
doc.add_paragraph('Flask Dev Server: Single-threaded, auto-reloads on code changes', style='List Bullet')
doc.add_paragraph('Port 5000: Default Flask port', style='List Bullet')

doc.add_heading('8.2 Running with Docker (Production-like)', level=2)
doc.add_paragraph('docker-compose up --build', style='No Spacing')
doc.add_paragraph('Gunicorn: Production WSGI server (multiple workers)', style='List Bullet')
doc.add_paragraph('Port 80: Standard HTTP port', style='List Bullet')
doc.add_paragraph('Volume: Database persists across container restarts', style='List Bullet')

doc.add_heading('8.3 Deployment to Heroku', level=2)
doc.add_paragraph('The project includes:', style='List Bullet')
doc.add_paragraph('Procfile: Tells Heroku how to run the app', style='List Bullet 2')
doc.add_paragraph('heroku.yml: Additional Heroku configuration', style='List Bullet 2')
doc.add_paragraph('GitHub Actions: CI/CD pipeline', style='List Bullet 2')

# PART 9
doc.add_heading('PART 9: COMMON INTERVIEW TOPICS', level=1)

doc.add_heading('9.1 How would you scale this?', level=2)
doc.add_paragraph('Use PostgreSQL instead of SQLite (multi-user support)', style='List Bullet')
doc.add_paragraph('Add user authentication (login/register)', style='List Bullet')
doc.add_paragraph('Use Kubernetes for orchestration', style='List Bullet')
doc.add_paragraph('Add caching (Redis)', style='List Bullet')
doc.add_paragraph('Database migrations (Alembic)', style='List Bullet')

doc.add_heading('9.2 What about security?', level=2)
doc.add_paragraph('Input validation (currently basic)', style='List Bullet')
doc.add_paragraph('CSRF protection (Flask-WTF)', style='List Bullet')
doc.add_paragraph('User authentication', style='List Bullet')
doc.add_paragraph('SQL injection prevention (ORM handles this)', style='List Bullet')
doc.add_paragraph('Environment variables for secrets', style='List Bullet')

doc.add_heading('9.3 Error Handling Improvements', level=2)
doc.add_paragraph('Current: Generic try/except blocks', style='List Bullet')
doc.add_paragraph('Better: Specific exceptions, logging, user-friendly messages', style='List Bullet')

doc.add_heading('9.4 Testing', level=2)
doc.add_paragraph('Unit tests for routes', style='List Bullet')
doc.add_paragraph('Integration tests for database', style='List Bullet')
doc.add_paragraph('Use pytest and Flask testing utilities', style='List Bullet')

# PART 10
doc.add_heading('PART 10: KEY PYTHON CONCEPTS USED', level=1)
table3 = doc.add_table(rows=8, cols=3)
table3.style = 'Light Grid Accent 1'
hdr_cells3 = table3.rows[0].cells
hdr_cells3[0].text = 'Concept'
hdr_cells3[1].text = 'Used In'
hdr_cells3[2].text = 'Purpose'

concepts_data = [
    ('Decorators', '@app.route()', 'Route definitions'),
    ('ORM Classes', 'Todo(db.Model)', 'Database abstraction'),
    ('Context Managers', 'with app.app_context()', 'Database initialization'),
    ('String Formatting', 'f-strings', 'URL building'),
    ('Boolean Logic', 'not task.complete', 'State toggling'),
    ('Exception Handling', 'try/except', 'Error management'),
    ('Environment Variables', 'os.environ.get()', 'Configuration'),
]

for i, (concept, used_in, purpose) in enumerate(concepts_data, 1):
    row_cells = table3.rows[i].cells
    row_cells[0].text = concept
    row_cells[1].text = used_in
    row_cells[2].text = purpose

# PART 11
doc.add_heading('PART 11: STUDY CHECKLIST', level=1)
doc.add_paragraph('Understand Flask request/response cycle', style='List Bullet')
doc.add_paragraph('Learn SQLAlchemy ORM query methods', style='List Bullet')
doc.add_paragraph('Practice Jinja2 template syntax', style='List Bullet')
doc.add_paragraph('Study Docker build/run process', style='List Bullet')
doc.add_paragraph('Understand Docker volumes for persistence', style='List Bullet')
doc.add_paragraph('Learn how URL routes map to functions', style='List Bullet')
doc.add_paragraph('Understand form submission (POST vs GET)', style='List Bullet')
doc.add_paragraph('Study database transaction concept (commit/rollback)', style='List Bullet')
doc.add_paragraph('Review error handling patterns', style='List Bullet')
doc.add_paragraph('Think about scalability improvements', style='List Bullet')

# Save the document
output_path = 'Todo_List_Flask_Study_Guide.docx'
doc.save(output_path)
print(f"✓ Document created successfully: {output_path}")
