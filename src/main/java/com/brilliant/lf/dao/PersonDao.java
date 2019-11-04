package com.brilliant.lf.dao;

import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.List;
import com.brilliant.lf.bean.Person;

@Mapper
public interface PersonDao {
    int insert(@Param("pojo") Person pojo);

    int insertSelective(@Param("pojo") Person pojo);

    int insertList(@Param("pojos") List<Person> pojo);

    int update(@Param("pojo") Person pojo);

    List<Person> getAll();
}
