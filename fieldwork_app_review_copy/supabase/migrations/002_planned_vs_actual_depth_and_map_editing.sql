alter table map_items
    add column if not exists planned_amount double precision default 0;

alter table map_items
    add column if not exists depth_m double precision default 0;

update map_items
set planned_amount = depth_m
where coalesce(planned_amount, 0) = 0
  and coalesce(depth_m, 0) > 0;

create index if not exists idx_map_items_project_type_item
    on map_items(project_id, item_type, item_id);
