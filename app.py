@admin_required
def admin_edit_center(center_id):
    c = db.get_or_404(Center, center_id); name = request.form.get('name','').strip()
    if name: c.name = name
    c.sort_order = request.form.get('sort_order', c.sort_order, type=int)
    log_action('center_updated', c.name); db.session.commit(); flash('Muudatused salvestatud.', 'success')
    return redirect(url_for('admin_home'))

@app.route('/admin/center/<int:center_id>/archive', methods=['POST'])
@admin_required
def admin_archive_center(center_id):
    c = db.get_or_404(Center, center_id); c.active = not c.active
    log_action('center_active_changed', f'{c.name}: active={c.active}'); db.session.commit()
    flash('Keskuse olek muudetud.', 'success'); return redirect(url_for('admin_home'))

@app.route('/admin/category', methods=['POST'])
@admin_required
def admin_add_category():
    name = request.form.get('name','').strip(); center_id = request.form.get('center_id', type=int)
    center = db.session.get(Center, center_id) if center_id else None
    if not name or not center: flash('Vali kaubanduskeskus ja sisesta kategooria nimi.', 'error')
    else:
        db.session.add(Category(center_id=center.id, name=name, sort_order=request.form.get('sort_order', 0, type=int)))
        log_action('category_created', f'{center.name} / {name}'); db.session.commit(); flash('Kategooria lisatud.', 'success')
    return redirect(url_for('admin_home'))

@app.route('/admin/category/<int:category_id>/edit', methods=['POST'])
@admin_required
def admin_edit_category(category_id):
    c = db.get_or_404(Category, category_id); name = request.form.get('name','').strip()
    if name: c.name = name
    c.sort_order = request.form.get('sort_order', c.sort_order, type=int)
    log_action('category_updated', c.name); db.session.commit(); flash('Kategooria uuendatud.', 'success'); return redirect(url_for('admin_home'))

@app.route('/admin/category/<int:category_id>/archive', methods=['POST'])
@admin_required
def admin_archive_category(category_id):
    c = db.get_or_404(Category, category_id); c.active = not c.active
    log_action('category_active_changed', f'{c.name}: active={c.active}'); db.session.commit(); flash('Kategooria olek muudetud.', 'success'); return redirect(url_for('admin_home'))

@app.route('/admin/meter', methods=['POST'])
@admin_required
def admin_add_meter():
    category_id = request.form.get('category_id', type=int); category = db.session.get(Category, category_id) if category_id else None
    name = request.form.get('name','').strip(); resource = request.form.get('resource_type','Muu'); unit = request.form.get('unit','').strip()
    allowed = {'Elekter','Külm vesi','Soe vesi','Gaas','Muu'}
    if not category or not name or resource not in allowed or not unit:
        flash('Kontrolli kategooriat, mõõdiku nime, ressurssi ja ühikut.', 'error')
    else:
        m = Meter(category_id=category.id, name=name, meter_code=request.form.get('meter_code','').strip() or None, resource_type=resource, unit=unit, sort_order=request.form.get('sort_order', 0, type=int))
        db.session.add(m); log_action('meter_created', f'{category.center.name} / {category.name} / {name}'); db.session.commit(); flash('Mõõdik lisatud.', 'success')
    return redirect(url_for('admin_home'))

@app.route('/admin/meter/<int:meter_id>/edit', methods=['POST'])
@admin_required
def admin_edit_meter(meter_id):
    m = db.get_or_404(Meter, meter_id); name = request.form.get('name','').strip(); unit = request.form.get('unit','').strip()
    resource = request.form.get('resource_type','Muu')
    if name and unit and resource in {'Elekter','Külm vesi','Soe vesi','Gaas','Muu'}:
        m.name=name; m.unit=unit; m.resource_type=resource; m.meter_code=request.form.get('meter_code','').strip() or None; m.sort_order=request.form.get('sort_order', m.sort_order, type=int)
        log_action('meter_updated', m.name); db.session.commit(); flash('Mõõdik uuendatud.', 'success')
    else: flash('Kontrolli mõõdiku andmeid.', 'error')
    return redirect(url_for('admin_home'))

@app.route('/admin/meter/<int:meter_id>/archive', methods=['POST'])
@admin_required
def admin_archive_meter(meter_id):
    m = db.get_or_404(Meter, meter_id); m.active = not m.active
    log_action('meter_active_changed', f'{m.name}: active={m.active}'); db.session.commit(); flash('Mõõdiku olek muudetud.', 'success'); return redirect(url_for('admin_home'))

@app.route('/admin/user', methods=['POST'])
@admin_required
def admin_add_user():
    username=request.form.get('username','').strip(); password=request.form.get('password',''); role=request.form.get('role','user')
    if not username or len(password) < 10 or role not in {'user','admin'}:
        flash('Kasutajanimi on kohustuslik ja parool peab olema vähemalt 10 märki.', 'error')
    elif User.query.filter_by(username=username).first(): flash('See kasutajanimi on juba kasutusel.', 'error')
    else:
        u=User(username=username, role=role); u.set_password(password); db.session.add(u); log_action('user_created', f'{username} ({role})'); db.session.commit(); flash('Kasutaja lisatud.', 'success')
    return redirect(url_for('admin_home'))

@app.route('/admin/user/<int:user_id>/toggle', methods=['POST'])
@admin_required
def admin_toggle_user(user_id):
    u=db.get_or_404(User,user_id)
    if u.id == current_user().id: flash('Enda kontot ei saa siin deaktiveerida.', 'error')
    else:
        u.active=not u.active; log_action('user_active_changed', f'{u.username}: active={u.active}'); db.session.commit(); flash('Kasutaja olek muudetud.', 'success')
    return redirect(url_for('admin_home'))

@app.route('/admin/user/<int:user_id>/password', methods=['POST'])
@admin_required
def admin_reset_password(user_id):
    u=db.get_or_404(User,user_id); password=request.form.get('password','')
    if len(password) < 10: flash('Parool peab olema vähemalt 10 märki.', 'error')
    else: u.set_password(password); log_action('password_reset', u.username); db.session.commit(); flash('Parool muudetud.', 'success')
    return redirect(url_for('admin_home'))

@app.errorhandler(403)
def forbidden(_): return render_template('error.html', code=403, message='Sul puudub sellele lehele ligipääs.'), 403
@app.errorhandler(404)
def not_found(_): return render_template('error.html', code=404, message='Lehte ei leitud.'), 404

with app.app_context():
    db.create_all()
    # Optional first-run admin bootstrap. Never hardcode production credentials.
    bootstrap_user = os.environ.get('BOOTSTRAP_ADMIN_USER')
    bootstrap_password = os.environ.get('BOOTSTRAP_ADMIN_PASSWORD')
    if bootstrap_user and bootstrap_password and not User.query.filter_by(role='admin').first():
        admin = User(username=bootstrap_user, role='admin'); admin.set_password(bootstrap_password); db.session.add(admin); db.session.commit()

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=False)
